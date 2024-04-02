"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");
var ExcelJS = require("exceljs");
var summary_json = JSON.parse(fs.readFileSync('report/report_summary.json', 'utf8'));
var result = [];
function getCloudSuitabilityTxt(cloudSuitabilityValue) {
    if (cloudSuitabilityValue >= 0 && cloudSuitabilityValue <= 49) {
        return "Low (Refactor)";
    }
    else if (cloudSuitabilityValue >= 50 && cloudSuitabilityValue <= 69) {
        return "Medium (Replatform)";
    }
    else if (cloudSuitabilityValue >= 70 && cloudSuitabilityValue <= 100) {
        return "High (Rehost)";
    }
    else {
        return "N/A";
    }
}
function getMigrationComplexityTxt(migrationComplexityType) {
    if (migrationComplexityType = 1) {
        return "Low";
    }
    else if (migrationComplexityType = 2) {
        return "Medium";
    }
    else if (migrationComplexityType = 3) {
        return "High";
    }
    else {
        return "N/A";
    }
}
console.log('Processing Data...');
summary_json.appDetails.forEach(function (app) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p;
    var detailed_json = JSON.parse(fs.readFileSync("report/report_app_".concat(app.appName, ".json"), 'utf8'));
    result.push({
        project: 'Dealer Service Agenda System',
        appName: app === null || app === void 0 ? void 0 : app.appName,
        cloudSuitabilityType: app === null || app === void 0 ? void 0 : app.cloudSuitabilityType,
        cloudSuitabilityValue: app === null || app === void 0 ? void 0 : app.cloudSuitabilityValue,
        cloudSuitabilityText: getCloudSuitabilityTxt(app === null || app === void 0 ? void 0 : app.cloudSuitabilityValue),
        totalAntiPatternCount: app === null || app === void 0 ? void 0 : app.totalAntiPatternCount,
        migrationComplexityType: app === null || app === void 0 ? void 0 : app.migrationComplexityType,
        migrationComplexityText: getMigrationComplexityTxt(app === null || app === void 0 ? void 0 : app.migrationComplexityType),
        migrationRecommendation: app === null || app === void 0 ? void 0 : app.migrationRecommendation,
        primaryFramework: (_a = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _a === void 0 ? void 0 : _a.primaryFramework,
        frontEndFramework: (_b = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _b === void 0 ? void 0 : _b.frontEndFramework,
        guiLanguage: (_c = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _c === void 0 ? void 0 : _c.guiLanguage,
        primaryLanguage: (_d = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _d === void 0 ? void 0 : _d.primaryLanguage,
        database: (_e = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _e === void 0 ? void 0 : _e.database,
        appServer: (_f = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _f === void 0 ? void 0 : _f.appServer,
        authentication: (_g = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _g === void 0 ? void 0 : _g.authentication,
        buildTool: (_h = detailed_json.techChar) === null || _h === void 0 ? void 0 : _h.buildTool,
        externalInterfaceList: (_j = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.externalInterfaceList) === null || _j === void 0 ? void 0 : _j.mapCount,
        externalInterfaceListText: ((_k = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.externalInterfaceList) === null || _k === void 0 ? void 0 : _k.mapCount) ? Object.entries((_l = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.externalInterfaceList) === null || _l === void 0 ? void 0 : _l.mapCount).map(function (_a) {
            var k = _a[0], v = _a[1];
            return "".concat(k, ": ").concat(v);
        }).join("\r\n") : "N/A",
        backingServices: (_m = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _m === void 0 ? void 0 : _m.backingServices,
        backingServicesText: ((_o = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _o === void 0 ? void 0 : _o.backingServices) ? Object.entries((_p = detailed_json === null || detailed_json === void 0 ? void 0 : detailed_json.techChar) === null || _p === void 0 ? void 0 : _p.backingServices).map(function (_a) {
            var k = _a[0], v = _a[1];
            return "".concat(k, ": ").concat(v);
        }).join("\r\n") : "N/A",
    });
});
var rows = result.map(function (row) {
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
    };
});
console.log('Preparing Excel Sheet...');
// Create a new Excel workbook
var workbook = new ExcelJS.Workbook();
// Add a worksheet to the workbook
var worksheet = workbook.addWorksheet('Assessment Data');
// Define headers
var headers = Object.keys(rows[0]);
var headerRow = worksheet.addRow(headers);
headerRow.height = 100;
// Set fill color for header cells
headerRow.eachCell(function (cell, index) {
    cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '00008B' } // Yellow color
    };
    cell.font = { color: { argb: 'FFFFFF' } }; // Red font color
});
var columnsToWrap = ['External Interfaces (Interface Name with Count)', 'Backing Services'];
rows.forEach(function (row, index) {
    var values = headers.map(function (header) { return row[header]; });
    var worksheetRow = worksheet.addRow(values);
    // Set wrap text for specific columns
    columnsToWrap.forEach(function (columnName) {
        var columnIndex = headers.indexOf(columnName);
        var cell = worksheetRow.getCell(columnIndex + 1); // Add 1 to columnIndex to account for header row
        cell.alignment = { wrapText: true };
    });
});
// Adjust column widths based on content length
worksheet.columns.forEach(function (column) {
    var maxLength = 0;
    column.eachCell({ includeEmpty: true }, function (cell) {
        var columnWidth = cell.value ? cell.value.toString().length : 0;
        maxLength = Math.max(maxLength, columnWidth);
    });
    column.width = maxLength < 10 ? 10 : maxLength; // Minimum width of 10 characters
});
// Adjust row heights based on content length
worksheet.eachRow(function (row, rowNumber) {
    var maxHeight = 0;
    row.eachCell({ includeEmpty: true }, function (cell) {
        var lines = cell.value ? cell.value.toString().split('\n') : [''];
        var cellHeight = lines.length;
        maxHeight = Math.max(maxHeight, cellHeight);
    });
    row.height = maxHeight * 15; // Assuming each line height is 15 points
});
var filePath = 'Assessment Report.xlsx';
(function () { return __awaiter(void 0, void 0, void 0, function () {
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0: return [4 /*yield*/, workbook.xlsx.writeFile(filePath)];
            case 1:
                _a.sent();
                console.log('Report Generated Successfully...Assessment Report.xlsx');
                return [2 /*return*/];
        }
    });
}); })();
