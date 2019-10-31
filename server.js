var express = require("express");
var app = express();
var bodyParser = require("body-parser");

var aspose = aspose || {};
aspose.cells = require("aspose.cells");

app.use(bodyParser.urlencoded({extended: false}));

app.post('/calc', function (req, res) {
    var wb = new aspose.cells.Workbook("Aspose.-.Dynamic.DCF.xlsx");
    var cells = wb.getWorksheets().get("Sheet1").getCells();
    cells.get("C2").putValue(req.body.wacc, true);
    cells.get("C3").putValue(req.body.grouth_rate, true);
    cells.get("C4").putValue(req.body.debt, true);
    cells.get("C5").putValue(req.body.cce, true);
    wb.calculateFormula();
    r = {}
    r["c13"] = cells.get("c13").getStringValue();
    r["c14"] = cells.get("c14").getStringValue();
    r["c15"] = cells.get("c15").getStringValue();
    r["c16"] = cells.get("c16").getStringValue();
    r["c17"] = cells.get("c17").getStringValue();

    res.end(JSON.stringify(r))
 })

app.use(express.static(".")).listen(8080);