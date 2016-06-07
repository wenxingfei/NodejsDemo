var express = require("express");
var wrap = require("co-express");
var xlsx1 = require('node-xlsx');
var fs = require('fs');
var excel = require('excel-export');
var XLSX = require('xlsx');
var xl = require('node-xlrd');
var xl2 = require('../node_modules/xlsx-master/index');
var ejsExcel = require("ejsexcel");
var exlBuf = fs.readFileSync("./modules/template2.xlsx");
var db = require("./db");
var ObjectId = require("mongodb").ObjectID;
var _ =require("lodash");
var router = express.Router();

var users=db.collection("users");

router.get('/',wrap(function* (req,res,next){
    var userDatas= yield users.find().toArrayAsync();
    _.each(userDatas,function(row){
        row.id=row._id;
        delete row._id;
    });
    res.send(userDatas);
}));

router.post('/',wrap(function* (req,res,next){
    var user= req.body;
    var rs= yield users.insertAsync(user);
    var id= rs.insertedIds[0];
    res.send({newid:id});
}));

router.put('/:id',wrap(function* (req,res,next){
    var id=req.param("id");
    var oid=new ObjectId(id);
    var user= req.body;
    delete user.id;
    var rs= yield users.updateAsync({_id:oid},user);
    res.send({});
}));

router.delete('/:id',wrap(function* (req,res,next){
    var id=req.param("id");
    var oid=new ObjectId(id);
    var rs= yield users.deleteOneAsync({_id:oid});
    res.send({});
}));

router.get('/export1', wrap(function* (req, res,next){
    //var obj = {"worksheets":[{"data":[["姓名","性别","年龄"],["李晓龙","男","24"]]}]};
    //var file = xlsx.build(obj);
    //var data = [[1,2,3],[true, false, null, 'sheetjs'],['foo','bar',new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];
    var userDatas= yield users.find().toArrayAsync();
    var data = [];
    var pros = [];
    if(userDatas.length){
        for(var pro in userDatas[0]){
            pros.push(pro);
        }
    }
    data.push(pros);
    for(var i = 0; i < userDatas.length; i++){
        var row = [];
        for(var j = 0; j < pros.length; j++){
            row.push(userDatas[i][pros[j]])
        }
        data.push(row);
    }
    var box = xlsx1.parse('./modules/template2.xlsx');
    //box[0].data = data;
    //var buffer = xlsx1.build([{name: "mySheetName", data: data}]);
    var buffer = xlsx1.build(box);
    //fs.writeFileSync('./public/temp/user.xlsx', buffer, 'binary');
    //res.send('user.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
    res.end(buffer, 'binary');
}));

router.get('/export2', wrap(function* (req, res,next){
    var conf = {};
    conf.name = "mysheet1";
    conf.cols = [
        {caption:'string', type:'string', captionStyleIndex: 1, width:28},
        {caption:'date', type:'date', },
        {caption:'bool', type:'bool', },
        {caption:'number', type:'number', }
    ];
    conf.rows = [
        ['pi', (new Date(2013, 4, 1)).getJulian(), true, 3.14],
        ["e", (new Date(2012, 4, 1)).getJulian(), false, 2.7182]
    ];
    //conf.stylesXmlFile = "styles.xml";
    /*conf.name = "mysheet";
    conf.cols = [{
        caption:'string',
        type:'string',
        beforeCellWrite:function(row, cellData){
            return cellData.toUpperCase();
        },
        width:28
    },{
        caption:'date',
        type:'date',
        beforeCellWrite:function(){
            var originDate = new Date(Date.UTC(1899,11,30));
            return function(row, cellData, eOpt){
                if (eOpt.rowNum%2){
                    eOpt.styleIndex = 1;
                }
                else{
                    eOpt.styleIndex = 2;
                }
                if (cellData === null){
                    eOpt.cellType = 'string';
                    return 'N/A';
                } else
                    return (cellData - originDate) / (24 * 60 * 60 * 1000);
            }
        }()
    },{
        caption:'bool',
        type:'bool'
    },{
        caption:'number',
        type:'number'
    }];
    conf.rows = [
        ['pi', new Date(Date.UTC(2013, 4, 1)), true, 3.14],
        ["e", new Date(2012, 4, 1), false, 2.7182],
        ["M&M<>'", new Date(Date.UTC(2013, 6, 9)), false, 1.61803],
        ["null date", null, true, 1.414]
    ];*/
    var result = excel.execute(conf);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
    res.end(result, 'binary');
}));

router.get('/export3', wrap(function* (req, res,next){
    var userDatas= yield users.find().toArrayAsync();
    var data = [];
    var pros = [];
    if(userDatas.length){
        for(var pro in userDatas[0]){
            pros.push(pro);
        }
    }
    data.push(pros);
    for(var i = 0; i < userDatas.length; i++){
        var row = [];
        for(var j = 0; j < pros.length; j++){
            row.push(userDatas[i][pros[j]])
        }
        data.push(row);
    }
    var ws = {
        s:{
            "!row" : [{wpx: 67}]
        }
    };
    ws['!cols']= [];
    for(var n = 0; n != data[0].length; ++n){
        ws['!cols'].push({
            wpx: 170
        });
    }
    var range = {
        s : {
            c : 10000000,
            r : 10000000,
        },
        e : {
            c : 0,
            r : 0
        }
    };
    for (var R = 0; R != data.length; ++R) {
        for (var C = 0; C != data[R].length; ++C) {
            if (range.s.r > R)
                range.s.r = R;
            if (range.s.c > C)
                range.s.c = C;
            if (range.e.r < R)
                range.e.r = R;
            if (range.e.c < C)
                range.e.c = C;
            var cell = {
                v : data[R][C],
                s:{
                    fill: { fgColor: { rgb: "#FF0000"}},
                    alignment: {horizontal: "center" ,vertical: "center"},
                }
            };
            if (cell.v == null)
                continue;
            var cell_ref = XLSX.utils.encode_cell({
                c : C,
                r : R
            });

            if ( typeof cell.v === 'number')
                cell.t = 'n';
            else if ( typeof cell.v === 'boolean')
                cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            } else
                cell.t = 's';
            if(R){
                delete cell.s.fill;
            }
            ws[cell_ref] = cell;
        }
    }
    data.fileName = "./public/temp/user.xlsx";
    var workbook = new Workbook();
    var wsName = data.fileName.split(".xlsx")[0];
    workbook.SheetNames.push(wsName);
    workbook.Sheets[wsName] = ws;
    if (range.s.c < 10000000)
        ws['!ref'] = XLSX.utils.encode_range(range);
    var wopts = {
        bookType : 'xlsx',
        bookSST : false,
        type : 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    XLSX.writeFile(workbook, data.fileName);
    //return wbout;
    res.send("user.xlsx");
}));

router.get('/export4', wrap(function* (req, res, next){
    var data = [
        [{"table_name":"现金报表","date": '2014-04-09'}],
        [ { "cb1":"001","cb1_":"002","bn1":"1","bn1_":"1","cn1":"1","cn1_":"1","num1":"1","num1_":"1",
            "cb5":"001","cb5_":"002","bn5":"1","bn5_":"1","cn5":"1","cn5_":"1","num1":"1","num5_":"1",
            "cb10":"001","cb10_":"002","bn10":"1","bn10_":"1","cn10":"1","cn10_":"1","num10":"1","num10_":"1",
            "cb20":"001","cb20_":"002","bn20":"1","bn20_":"1","cn20":"1","cn20_":"1","num20":"1","num20_":"1",
            "cb50":"001","cb50_":"002","bn50":"1","bn50_":"1","cn50":"1","cn50_":"1","num50":"1","num50_":"1",
            "cb100":"001","cb100_":"002","bn100":"1","bn100_":"1","cn100":"1","cn100_":"1","num100":"1","num100_":"1"
        },{ "cb1":"001","cb1_":"002","bn1":"1","bn1_":"1","cn1":"1","cn1_":"1","num1":"1","num1_":"1",
            "cb5":"001","cb5_":"002","bn5":"1","bn5_":"1","cn5":"1","cn5_":"1","num1":"1","num5_":"1",
            "cb10":"001","cb10_":"002","bn10":"1","bn10_":"1","cn10":"1","cn10_":"1","num10":"1","num10_":"1",
            "cb20":"001","cb20_":"002","bn20":"1","bn20_":"1","cn20":"1","cn20_":"1","num20":"1","num20_":"1",
            "cb50":"001","cb50_":"002","bn50":"1","bn50_":"1","cn50":"1","cn50_":"1","num50":"1","num50_":"1",
            "cb100":"001","cb100_":"002","bn100":"1","bn100_":"1","cn100":"1","cn100_":"1","num100":"1","num100_":"1"
        }]
    ];
    ejsExcel.renderExcelCb(exlBuf, data, function(exlBuf2){
        fs.writeFileSync("./public/temp/report1.xlsx", exlBuf2);
        res.send('report1.xlsx');
    });
}));

function arrays_equal(a,b) { return !(a<b || b<a); }
router.get('/export5', wrap(function* (req, res, next){
    var xls = new xl2();
    xls.loadFile('./modules/template2.xlsx');
    var sheet0 = xls.getSheet(0);
    var data = [];
    data.push([1, "张三", "zhangsan@huawei.com", "123"]);
    data.push([2, "李四", "李四@huawei.com", "456"]);
    for(var i = 0; i < data.length; i++){
        sheet0.write('A' + (i + 2), data[i]);
    }
    //xls.writeFile('./public/temp/test.xlsx');
    //res.send('test.xlsx')
    var buffer = xls.data();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
    res.end(buffer, 'binary');
}));

function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

module.exports = router;