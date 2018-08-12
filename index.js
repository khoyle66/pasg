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
    
    resetVariable();    
    var HTMLOUT = document.getElementById('htmlout');    


    //Get Spreadsheet
    var first_sheet_name = wb.SheetNames[0];
    sheetName = first_sheet_name;

    var htmlstr = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);

    var lab_hours = 0;
    var lab_cost = 0;
    var lab_sales = 0;
    var start_date;
    var end_date;
    var project_code =  htmlstr[1].__EMPTY_2;
    //console.log(project_code);

    //Load Project Data
    var projectFile = fs.readFileSync("projects.json");
    // Define to JSON type
    var projects = JSON.parse(projectFile);
   
    var project;
    for (var i = 0; i < projects.length; i++) {
        if( projects[i].Code==project_code) {
            project = projects[i];
        }
    }  
    console.log(project.Items[2]) ;

    if(project == undefined) {
        fin_sum.style.display = "none";           
    } 
    else {  

        var labourItem = project.Items[0];
        var subContractItem = project.Items[1];
        var expenseItem = project.Items[2];
        var capitalItem = project.Items[3];

        //Extract Data from Spreadsheet
        htmlstr.forEach(function(record) {
            
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
            if(isLabour(record)) {
                if(!isNaN(record.__EMPTY_8)) {
                    lab_hours += Number(record.__EMPTY_8);;
                }               
                if(!isNaN(record.__EMPTY_9)) {
                    labourItem.costActual += Number(record.__EMPTY_9);
                }
                if(isBillable(record) && !isNaN(record.__EMPTY_10)) {
                    labourItem.revActual += Number(record.__EMPTY_10);;
                }
            }
        });

        //labourItem.costActual = lab_cost;
        //labourItem.revActual = lab_sales;

        //Display Financial Summary Tables
        fin_sum.style.display = "block";   
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

        finItems.push(labourItem);
        finItems.push(subContractItem);
        finItems.push(expenseItem);
        finItems.push(capitalItem);

        //var totalItem = new FinancialItem("Total");

        addCostTable(table_cost, finItems);

        var totalRevItem = addRevenueTable(table_revenue, finItems); 
        //var tot_budget = lab_sales-lab_hours;
        var totalItem = addMarginTable(table_margin, finItems);

        addMarginTablePerc(table_margin, totalRevItem, totalItem);
    }
});  

function resetVariable() {
    htmlout.innerHTML = "";    
    P_Code.innerHTML = "";
    P_Name.innerHTML = "";
    PM_Name.innerHTML = "";
    Cust_Name.innerHTML = "";
    Sponsor.innerHTML = "";
    ContractBasis.innerHTML = "";
    Category.innerHTML = "";
    Stage.innerHTML = "";
    Start_Date.innerHTML = "";
    PSRFrequency.innerHTML = "";
    BACompDate.innerHTML = "";
    RepPeriodFrom.innerHTML = "";
    RepPeriodTo.innerHTML = "";
    CACompDate.innerHTML = "";
    FCompDate.innerHTML = "";
    PercComp.innerHTML = "";
    deletAllRows(table_cost);
    deletAllRows(table_revenue);
    deletAllRows(table_margin);
};

function deletAllRows(table) {
    for(var i = table.rows.length - 1; i > 0; i--)
    {
        table.deleteRow(i);
    }    
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
    addTableRow(table, totalItem.description, totalItem.costBudget, totalItem.costPCRs, totalItem.costActual, totalItem.costEtc); 
};
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
    addTableRow(table, totalItem.description, totalItem.revBudget, totalItem.revPCRs, totalItem.revActual, totalItem.revEtc) ;
    return totalItem;        
};
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
    return totalItem;
};

function addMarginTablePerc(table, totalRevenue, totalMargin) {
    var totalBudget = totalMargin.revBudget / totalRevenue.revBudget * 100;
    if(totalRevenue.revPCRs==0)
        totalPCRs = 0;
    else
        totalPCRs = totalMargin.revPCRs / totalRevenue.revPCRs * 100;
    var totalActual = totalMargin.revActual / totalRevenue.revActual * 100;
    var totalEtc = totalMargin.revEtc / totalRevenue.revEtc * 100;
    addTableRow(table, "Margin Total %", totalBudget,"",totalActual,totalEtc);
};

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
};
function displayDate(date) {
    return moment(date,'YYYY/MM/DD').format('DD-MMM-YY'); 
};


//-----------------------------------------------------------------------
// Spreadsheet Functions
//-----------------------------------------------------------------------

function processSpreasheet(htmlstr) {

};
function isLabour(value) {
    if(value.__EMPTY_6 == 'Labour' || value.__EMPTY_6 == 'Contract Labour') {
        return true;
    }
    return false;
};
function isBillable(value) {
    if(value.__EMPTY_11 == 'Billable') {
        return true;
    }
    return false;
};

