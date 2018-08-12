const electron = require('electron');
const {ipcRenderer} = electron;
const XLSX = require('xlsx');
const ul = document.querySelector('ul');    
const moment = require('moment');
const fs = require('fs');

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
    HTMLOUT.innerHTML = "";    

    //Display Financial Summary Tables
    fin_sum.style.display = "block";
    //Get Spreadsheet
    var first_sheet_name = wb.SheetNames[0];
    sheetName = first_sheet_name;

    var htmlstr = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);

    var lab_hours = 0;
    var lab_cost = 0;
    var lab_sales = 0;
    var start_date;
    var end_date;
    var project_code;
    //Extract Data from Spreadsheet
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
        project_code = record.__EMPTY_2; 
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

    //Load Project Data
    var projectFile = fs.readFileSync("projects.json");
    // Define to JSON type
    var projects = JSON.parse(projectFile);
    console.log(projects);
    console.log(projects[0]);
    var projects;
    for (var i = 0; i < projects.length; i++) {
        if( projects[i].Code==project_code) {
            project = projects[i];
        }
    }   
   
    //Set HTML Values
    P_Code.innerHTML = project_code;           
    P_Name.innerHTML = project.Name;    
    PM_Name.innerHTML = project.Manager;
    Cust_Name.innerHTML = project.CustomerName;
    Sponsor.innerHTML = project.Sponsor;
    ContractBasis.innerHTML = project.ContractBasis;
    Category.innerHTML = project.Category;
    Stage.innerHTML = project.Stage;
    Start_Date.innerHTML = displayDate(project.StartDate);
    PSRFrequency.innerHTML = project.PSRFrequency;
    BACompDate.innerHTML = displayDate(project.BaselineApprovedCompletionDate);
    RepPeriodFrom.innerHTML = displayDate(project.ReportingPeriodFrom);
    RepPeriodTo.innerHTML = displayDate(project.ReportingPeriodTo);
    CACompDate.innerHTML = displayDate(project.CurrentApprovedCompletionDate);
    FCompDate.innerHTML = displayDate(project.ForecastProjectCompletionDate);
    PercComp.innerHTML = project.OverallPercComplete + "%";

    HTMLOUT.innerHTML =  XLSX.utils.sheet_to_html(wb.Sheets[first_sheet_name],{editable:true});
    //Start_Date.innerHTML = start_date.format('DD-MMM-YY');
    //W_End.innerHTML = end_date.format('DD-MMM-YY'); 
    //TEMP FIELDS
    var COST_PCR = 0;
    var REV_PCR = 0;
    ///
    var finItems = new Array();
    
    var labourItem = new FinancialItem("ASG Labour");
    labourItem.costBudget = project.CostBudget;
    labourItem.costPCRs = COST_PCR;
    labourItem.costActual = lab_cost;
    labourItem.costEtc = project.CostBudget-lab_cost;
    labourItem.revBudget = project.PriceBudget;    
    labourItem.revPCRs = REV_PCR;
    labourItem.revActual = lab_sales;
    labourItem.revEtc = project.PriceBudget - lab_sales;
    var subContractItem = new FinancialItem("ASG Subcontractors");
    var expenseItem = new FinancialItem("Travel/other Expenses");
    var capitalItem = new FinancialItem("Capital expenditure");

    finItems.push(labourItem);
    finItems.push(subContractItem);
    finItems.push(expenseItem);
    finItems.push(capitalItem);

    //var totalItem = new FinancialItem("Total");

    addCostTable(table_cost, finItems);

    addRevenueTable(table_revenue, finItems); 
    //var tot_budget = lab_sales-lab_hours;
    addMarginTable(table_margin, finItems);

});  

function displayDate(date) {
    return moment(date,'YYYY/MM/DD').format('DD-MMM-YY'); 
};

function FinancialItem (description) { 
    this.description = description;
    this.costBudget = 0;
    this.costPCRs = 0;
    this.costActual = 0;
    this.costEtc = 0;    
    this.revBudget = 0;
    this.revPCRs = 0;
    this.revActual = 0;
    this.revEtc = 0;        
};

function addCostTable(table, items) {   
    var totalItem = new FinancialItem("Cost Total");  
    for (var i = 0, len = items.length; i < len; i++) {
        item = items[i];
        addTableRow(table, item.description, item.costBudget, item.costPCRs, item.costActual, item.costEtc)    
        totalItem.costBudget += item.costBudget;
        totalItem.costPCRs += item.costPCRs;
        totalItem.costActual += item.costActual;
        totalItem.costEtc += item.costEtc;
      }   
    addTableRow(table, totalItem.description, totalItem.costBudget, totalItem.costPCRs, totalItem.costActual, totalItem.costEtc)     
}
function addRevenueTable(table, items) {
    var totalItem = new FinancialItem("Price Total");  
    for (var i = 0, len = items.length; i < len; i++) {
        item = items[i];
        addTableRow(table, item.description, item.revBudget, item.revPCRs, item.revActual, item.revEtc)    
        totalItem.revBudget += item.revBudget;
        totalItem.revPCRs += item.revPCRs;
        totalItem.revActual += item.revActual;
        totalItem.revEtc += item.revEtc;
      }   
    addTableRow(table, totalItem.description, totalItem.revBudget, totalItem.revPCRs, totalItem.revActual, totalItem.revEtc)         
}
function addMarginTable(table, items) {
    var totalItem = new FinancialItem("Margin Total $");    
    for (var i = 0, len = items.length; i < len; i++) {
        item = items[i];
        var totalBudget = item.revBudget - item.costBudget;
        var totalPCRs = item.revPCRs - item.costPCRs;
        var totalActual = item.revActual - item.costActual;
        var totalEtc = item.revEtc - item.costEtc;
        addTableRow(table, item.description, totalBudget, totalPCRs, totalActual, totalEtc)    
        totalItem.revBudget += totalBudget;
        totalItem.revPCRs += totalPCRs;
        totalItem.revActual += totalActual;
        totalItem.revEtc += totalEtc;
      }   
    addTableRow(table, totalItem.description, totalItem.revBudget, totalItem.revPCRs, totalItem.revActual, totalItem.revEtc)   
    addTableRow(table, "Margin Total %", 0,0,0,0);
}

function addTableRow(table_name, description, initialBudget, approvedPCRs, actual, etc) {
    var row = table_name.insertRow();
    row.insertCell(0).innerHTML = description;
    row.insertCell(1).innerHTML = initialBudget.toLocaleString();
    row.insertCell(2).innerHTML = approvedPCRs.toLocaleString();
    var curBudget = initialBudget + approvedPCRs;
    row.insertCell(3).innerHTML = curBudget.toLocaleString();
    row.insertCell(4).innerHTML = actual.toLocaleString();
    row.insertCell(5).innerHTML = etc.toLocaleString();
    var eac = actual+etc;
    row.insertCell(6).innerHTML = eac.toLocaleString();
    var vac = eac - curBudget;
    row.insertCell(7).innerHTML = vac.toLocaleString();
    var vacPerc = 0;
    if(vac != 0) {
        vacPerc = (vac/curBudget)*100;
    }
    row.insertCell(8).innerHTML = vacPerc.toLocaleString();
}

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
