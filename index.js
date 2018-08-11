const electron = require('electron');
const {ipcRenderer} = electron;
const XLSX = require('xlsx');
const ul = document.querySelector('ul');    
const moment = require('moment');

// Add item
ipcRenderer.on('item:add', function(e, item){
    ul.className = 'collection';
    const li = document.createElement('li');
    li.className = 'collection-item';
    const itemText = document.createTextNode(item);
    li.appendChild(itemText);
    ul.appendChild(li);
});

//Clear item
ipcRenderer.on('item:clear', function(){
    ul.innerHTML = '';
    ul.className = '';  
});  

//Remove item
ul.addEventListener('dblclick', removeItem);

function removeItem(e){
    e.target.remove();
    if(ul.children.length == 0) {
        ul.className = '';                    
    }
}


//  Add document
ipcRenderer.on('doc:add', function(e, wb){
    var HTMLOUT = document.getElementById('htmlout');
    var LAB_HOURS = document.getElementById('Labour_Hours');
    var LAB_COST = document.getElementById('Labour_Cost');
    var LAB_SALES = document.getElementById('Labour_Sales');
    var first_sheet_name = wb.SheetNames[0];
    HTMLOUT.innerHTML = "";

    sheetName = first_sheet_name;

    // var htmlstr = XLSX.utils.sheet_to_html(wb.Sheets[sheetName],{editable:true});
    var htmlstr = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
    var lab_hours = 0;
    var lab_cost = 0;
    var lab_sales = 0;
    var start_date;
    var end_date;
    htmlstr.forEach(function(record) {
        //console.log(record);
        var tran_date = moment(record.__EMPTY,'DD/MM/YYYY'); 
        var period_ending = moment(record.__EMPTY_1,'DD/MM/YYYY');  
        if(tran_date.isValid()) {
            if (!start_date || tran_date < start_date) {                    
                    start_date = tran_date;                        
            }          
        }
        if(period_ending.isValid()) {
            if (!end_date || period_ending > end_date) {
                end_date = period_ending;
            }            
        }            
        P_Code.innerHTML = record.__EMPTY_2;           
        P_Name.innerHTML = record.__EMPTY_3;
        if(isLabour(record)) {
            if(!isNaN(record.__EMPTY_8)) {
                lab_hours += Number(record.__EMPTY_8);;
            }               
            if(!isNaN(record.__EMPTY_9)) {
                lab_cost += Number(record.__EMPTY_9);
            }
            if(isBillable(record) && !isNaN(record.__EMPTY_10)) {
                lab_sales += Number(record.__EMPTY_10);;
            }
        }
    });
    Start_Date.innerHTML = start_date.format('DD-MMM-YY');
    W_End.innerHTML = end_date.format('DD-MMM-YY');      
    LAB_HOURS.innerHTML = lab_hours;
    //LAB_COST.innerHTML = lab_cost.toLocaleString();
    LAB_SALES.innerHTML = lab_sales.toLocaleString();
    HTMLOUT.innerHTML =  XLSX.utils.sheet_to_html(wb.Sheets[first_sheet_name],{editable:true});

    //addTableRow(table_cost, createSummaryLine("ASG Subcontractors"));
    var lab_apcr = 0;
    var sub_apcr = 0;
    var exp_apcr = 0;
    var cap_apcr = 0;
    addTableRow(table_cost, "ASG Labour", 193590.80,lab_apcr,lab_cost,14812);
    addTableRow(table_cost, "ASG Subcontractors", 0,sub_apcr,0,0);
    addTableRow(table_cost, "Travel/other Expenses", 0,exp_apcr,0,0);
    addTableRow(table_cost, "Capital expenditure", 0,cap_apcr,0,0);
    var tot_apcr = lab_apcr + sub_apcr + exp_apcr + cap_apc;
    addTableRow(table_cost, "Cost Total", 193590.80,tot_apcr,lab_cost,14812);
});  

function isLabour(value) {
    if(value.__EMPTY_6 == 'Labour' || value.__EMPTY_6 == 'Contract Labour') {
        return true;
    }
    return false;
}
function isBillable(value) {
    if(value.__EMPTY_11 == 'Billable') {
        return true;
    }
    return false;
}

function addTableRow(table_name, description, initialBudget, approvedPCRs, actual, etc) {
    var row = table_name.insertRow();
    row.insertCell(0).innerHTML = description;
    row.insertCell(1).innerHTML = initialBudget.toLocaleString();
    row.insertCell(2).innerHTML = approvedPCRs.toLocaleString();
    var curBudget = initialBudget + approvedPCRs;
    row.insertCell(3).innerHTML = curBudget.toLocaleString();
    row.insertCell(4).innerHTML = actual.toLocaleString();
    etc= curBudget - actual;
    row.insertCell(5).innerHTML = etc.toLocaleString();
    var eac = actual+etc;
    row.insertCell(6).innerHTML = eac.toLocaleString();
    var vac = eac - curBudget;
    row.insertCell(7).innerHTML = vac.toLocaleString();
    var vacPerc = 0;
    if(vac != 0) {
        (vac/curBudget)*100;
    }
    row.insertCell(8).innerHTML = vacPerc.toLocaleString();
}