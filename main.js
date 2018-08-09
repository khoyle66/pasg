const electron = require('electron')
const {app, BrowserWindow, Menu, ipcMain} = require('electron')
const {dialog} = require('electron')
const fs = require('fs');
    // SET ENV
    //process.env.NODE_ENV = 'production';

  // Keep a global reference of the window object, if you don't, the window will
  // be closed automatically when the JavaScript object is garbage collected.
  let win
  let addWin

  function createWindow () {
    // Create the browser window.
    win = new BrowserWindow({width: 800, height: 600})
  
    // and load the index.html of the app.
    win.loadFile('index.html')
  
    // Open the DevTools.
    //win.webContents.openDevTools()
  
    // Build menu from template
    const mainMenu = Menu.buildFromTemplate(mainMenuTemplate);
    // Insert menu
    Menu.setApplicationMenu(mainMenu);

    // Emitted when the window is closed.
    win.on('closed', () => {
      // Dereference the window object, usually you would store windows
      // in an array if your app supports multi windows, this is the time
      // when you should delete the corresponding element.
      win = null;
      app.quit();
    })
  }
  
  // This method will be called when Electron has finished
  // initialization and is ready to create browser windows.
  // Some APIs can only be used after this event occurs.
  app.on('ready', createWindow)
  
  // Quit when all windows are closed.
  app.on('window-all-closed', () => {
    // On macOS it is common for applications and their menu bar
    // to stay active until the user quits explicitly with Cmd + Q
    if (process.platform !== 'darwin') {
      app.quit()
    }
  })
  
  app.on('activate', () => {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (win === null) {
      createWindow()
    }
  })
  
  // In this file you can include the rest of your app's specific main process
  // code. You can also put them in separate files and require them here.

    // Handle create add window
  
    function createAddWindow() {
        addWin = new BrowserWindow({
            width: 300, 
            height: 200,
            title:'Add Shopping List Item'
        })
    
        // and load the index.html of the app.
        addWin.loadFile('addWindow.html');

        // Garbage collection handle
        addWin.on('close', function(){
            addWin = null;
        });
    }

    //Catch item:add
    ipcMain.on('item:add', function(e, item) {
        console.log("Main:"+item);
        win.webContents.send('item:add',item);
        addWin.close();
    });

  //Create menu template
  const mainMenuTemplate = [
      {
          label:'File',
          submenu:[
              {
                  label: 'Add Items',
                  click(){
                      createAddWindow();
                  }
              },
              {
                  label: 'Open AX file',
                  click(){
                    dialog.showOpenDialog(
                        { 
                            filters: [
                                { name: 'all files', extensions: ['*']},
                                { name: 'json', extensions: ['json']},
                                { name: 'csv', extensions: ['csv'] }
                            ]
                        }, 
                        function (fileNames) {
                            if (fileNames === undefined) return;
                            var fileName = fileNames[0];
                            fs.readFile(fileName, 'utf-8', 
                                function (err, data) 
                                {
                                    console.log("Doc:"+data);
                                    //win.getElementById("editor").value = data;                     
                                    win.webContents.send('doc:add',data);
                                }
                            );                     
                        }
                    ); 
                  }
              },
              {
                label: 'Clear Items',
                click() {
                    win.webContents.send('item:clear');
                }
              },
              {
                  label:'Quit',
                  accelerator: process.platform == 'darwin' ?'Command+Q' : 'Ctrl+Q',
                  click(){
                      app.quit();
                  }
              }
          ]
      }
  ];

  // if mac, add empty object to menu
  if(process.platform == 'darwin') {
      mainMenuTemplate.unshift({});
  }

  // Add developer tools item if not in prod
  if(process.env.NODE_ENV !== 'production') {
      mainMenuTemplate.push({
          label:'Developer Tools',
          submenu:[
              {
                  label: 'Toggle DevTools',
                  accelerator: process.platform == 'darwin' ?'Command+I' : 'Ctrl+I',
                  click(item, focusedWindow) {
                      focusedWindow.toggleDevTools();
                  }
              },
              {
                  role: 'reload'
              }
          ]
      })
  }