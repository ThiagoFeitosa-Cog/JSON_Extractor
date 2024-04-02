import * as fs from 'fs';
import * as ExcelJS from 'exceljs';

let summary_json = JSON.parse(fs.readFileSync('report/report_summary.json', 'utf8'));

const result : any[] = [];

function getCloudSuitabilityTxt(cloudSuitabilityValue) {
    if(cloudSuitabilityValue >= 0 && cloudSuitabilityValue <= 49)
    {
        return "Low (Refactor)";
    }
    else if(cloudSuitabilityValue >= 50 && cloudSuitabilityValue <= 69)
    {
        return "Medium (Replatform)";
    }
    else if(cloudSuitabilityValue >= 70 && cloudSuitabilityValue <= 100)
    {
        return "High (Rehost)";
    }
    else{
        return "N/A";
    }
}

function getMigrationComplexityTxt(migrationComplexityType) {
    if(migrationComplexityType= 1)
    {
        return "Low";
    }
    else if(migrationComplexityType= 2)
    {
        return "Medium";
    }
    else if(migrationComplexityType= 3)
    {
        return "High";
    }
    else{
        return "N/A";
    }
}

console.log('Processing Data...');

summary_json.appDetails.forEach(app => {
    let detailed_json = JSON.parse(fs.readFileSync(`report/report_app_${app.appName}.json`, 'utf8'));
    result.push({
        project: 'Dealer Service Agenda System',
        appName: app?.appName,
        cloudSuitabilityType: app?.cloudSuitabilityType,
        cloudSuitabilityValue: app?.cloudSuitabilityValue,
        cloudSuitabilityText: getCloudSuitabilityTxt(app?.cloudSuitabilityValue),
        totalAntiPatternCount: app?.totalAntiPatternCount,
        migrationComplexityType: app?.migrationComplexityType,
        migrationComplexityText: getMigrationComplexityTxt(app?.migrationComplexityType),
        migrationRecommendation: app?.migrationRecommendation,
        primaryFramework: detailed_json?.techChar?.primaryFramework,
        frontEndFramework: detailed_json?.techChar?.frontEndFramework,
        guiLanguage: detailed_json?.techChar?.guiLanguage,
        primaryLanguage: detailed_json?.techChar?.primaryLanguage,
        database: detailed_json?.techChar?.database,
        appServer: detailed_json?.techChar?.appServer,
        authentication: detailed_json?.techChar?.authentication,
        buildTool: detailed_json.techChar?.buildTool,
        externalInterfaceList: detailed_json?.externalInterfaceList?.mapCount,
        externalInterfaceListText: detailed_json?.externalInterfaceList?.mapCount ? Object.entries(detailed_json?.externalInterfaceList?.mapCount).map(([k,v]) => `${k}: ${v}`).join("\r\n") : "N/A",
        backingServices: detailed_json?.techChar?.backingServices,
        backingServicesText: detailed_json?.techChar?.backingServices ? Object.entries(detailed_json?.techChar?.backingServices).map(([k,v]) => `${k}: ${v}`).join("\r\n") : "N/A",
    });
});

let rows = result.map(row => {
    return {
        'Application': row.project,
        'Application Name (Modules)': row.appName,
        'Tech Stack Suitability': row.cloudSuitabilityText,
        'Total Antipattens/Impediments Count': row.totalAntiPatternCount,
        'Migration Complexity': row.migrationComplexityText,
        'Migration Recommendation': row.migrationRecommendation,
        'Web Frameworks': row.primaryFramework,
        'Frontend Framework': row.frontEndFramework,
        'GUI Language': row.guiLanguage,
        'Programming Language': row.primaryLanguage,
        'Database': row.database,
        'Web/Application Server': row.appServer,
        'Authentication': row.authentication,
        'Build Tool': row.buildTool,
        'External Interfaces (Interface Name with Count)': row.externalInterfaceListText,
        'Backing Services': row.backingServicesText
    }
});

console.log('Preparing Excel Sheet...');

// Create a new Excel workbook
const workbook = new ExcelJS.Workbook();

// Add a worksheet to the workbook
const worksheet = workbook.addWorksheet('Assessment Data');

// Define headers
const headers = Object.keys(rows[0]);
const headerRow = worksheet.addRow(headers);

headerRow.height = 100;

// Set fill color for header cells
headerRow.eachCell((cell, index) => {
    cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '00008B' } // Yellow color
    };
    cell.font = { color: { argb: 'FFFFFF' } }; // Red font color
});

const columnsToWrap = ['External Interfaces (Interface Name with Count)', 'Backing Services'];

rows.forEach((row, index) => {
    const values = headers.map(header => row[header]);
    const worksheetRow = worksheet.addRow(values);
    // Set wrap text for specific columns
    columnsToWrap.forEach(columnName => {
        const columnIndex = headers.indexOf(columnName);
        const cell = worksheetRow.getCell(columnIndex + 1); // Add 1 to columnIndex to account for header row
        cell.alignment = { wrapText: true };
    });
});

// Adjust column widths based on content length
worksheet.columns.forEach(column => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, cell => {
        const columnWidth = cell.value ? cell.value.toString().length : 0;
        maxLength = Math.max(maxLength, columnWidth);
    });
    column.width = maxLength < 10 ? 10 : maxLength; // Minimum width of 10 characters
});

// Adjust row heights based on content length
worksheet.eachRow((row, rowNumber) => {
    let maxHeight = 0;
    row.eachCell({ includeEmpty: true }, cell => {
        const lines = cell.value ? cell.value.toString().split('\n') : [''];
        const cellHeight = lines.length;
        maxHeight = Math.max(maxHeight, cellHeight);
    });
    row.height = maxHeight * 15; // Assuming each line height is 15 points
});

const filePath = 'Assessment Report.xlsx';

(async () => {
    await workbook.xlsx.writeFile(filePath);
    console.log('Report Generated Successfully...Assessment Report.xlsx');
  })();