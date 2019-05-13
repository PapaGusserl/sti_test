var express = require('express');
var fs = require('fs');
var Stimulsoft = require("stimulsoft-reports-js");
var bodyParser = require('body-parser');
var bcrypt = require('bcryptjs');
Stimulsoft.Base.StiFontCollection.addOpentypeFontFile("Roboto-Black.ttf");
var app = express();

app.use(bodyParser.json({limit: '10mb', extended: true}))
app.use(bodyParser.urlencoded({limit: '10mb', extended: true}))

app.post('/render', function(req, res) {
  console.log("POST /render " + JSON.stringify(req.body));
  var resp = render(req.body.report, req.body.data, req.body.into);
  res.writeHeader(resp.status, {"Content-Type": "application/json"});
  var body = {status: resp.text_status, body: resp.body};
  res.end(JSON.stringify(body));
})

app.listen(3000, function() {
    console.log("Static file server running at http://localhost:" + 3000 + "/\nCTRL + C to shutdown");
})

function render(report_data, data, format) {
    if(!report_data || !data) {
        return {
            status: 400,
            text_status: "error",
            body: "bad arguments"
        };
    }
    var report = new Stimulsoft.Report.StiReport();
    try {
        JSON.parse(report_data);
        report.load(report_data);
    } catch (e) {
        try {
            report.loadFile(report_data);
        } catch (e) {
            return {
                status: 400,
                text_status: "error",
                body: "cant load report"
            }
        }
    }
    report.dictionary.databases.clear();
    reg_dataset(report, data);
    reg_variables(report, data);
    report.render();
    var result;
    switch(format) {
        case 'html':
            return {
                status: 200,
                text_status: "success",
                body: report.exportDocument(Stimulsoft.Report.StiExportFormat.Html)
            }
            break;
        case 'pdf':
            return {
                status: 200,
                text_status: "success",
                body: report.exportDocument(Stimulsoft.Report.StiExportFormat.Pdf)
            }
            break;
        case 'xlsx':
            var format = Stimulsoft.Report.StiExportFormat.Excel2007;
            var mimeType = 'application/ms-excel';
            var extension = 'xlsx';
            var service = new Stimulsoft.Report.Export.StiExcel2007ExportService();
            var settings = new Stimulsoft.Report.Export.StiExcel2007ExportSettings();
            settings.excelType = Stimulsoft.Report.Export.StiExcelType.ExcelXml;
            var stream = new Stimulsoft.System.IO.MemoryStream();
            service.exportTo(report, stream, settings);
            return {
                status: 200,
                text_status: "success",
                body: Buffer.from(stream.toArray())
  			}
            break;
        case 'doc':
            var format = Stimulsoft.Report.StiExportFormat.Word2007;
            var mimeType = 'application/msword';
            var extension = 'doc';
            var service = new Stimulsoft.Report.Export.StiWord2007ExportService();
            var settings = new Stimulsoft.Report.Export.StiWord2007ExportSettings();
            var stream = new Stimulsoft.System.IO.MemoryStream();
            service.exportTo(report, stream, settings);
            return {
                status: 200,
                text_status: "success",
                body: Buffer.from(stream.toArray())
            }
            break;
        default:
            return {
                status: 400,
                text_status: "error",
                body: "no output format"
            };
            break;
    }
}

function reg_dataset(report, data) {
    if(!data["dataset"]) {
        return;
    }
    var datasets = data["dataset"];
    Object.keys(datasets).forEach(function(name) {
        var set = datasets[name];
        var dataSet = new Stimulsoft.System.Data.DataSet(name);
        dataSet.readJson(set);
        report.regData(name, name, dataSet);
    })
}

function reg_variables(report, data) {
    if(!data["variables"]) {
        return;
    }
    var variables = data["variables"];
    Object.keys(variables).forEach(function(name) {
        var value = variables[name];
        var variable = report.dictionary.variables.getByName(name);
        if (variable) {
            variable.value = value;
        }
    })
}
