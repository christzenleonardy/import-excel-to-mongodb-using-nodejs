'use strict';

// For import excel to mongo
const excelToJson = require('convert-excel-to-json');
var MongoClient = require('mongodb').MongoClient;
var url = "mongodb://localhost:27017/";

// For upload file
var express = require('express'); 
var app = express(); 
var bodyParser = require('body-parser');
var multer = require('multer');

app.use(bodyParser.json());

var storage = multer.diskStorage({ //multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './uploads/')
    },
    filename: function (req, file, cb) {
        var datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length -1])
    }
});

var upload = multer({ //multer settings
                storage: storage,
                fileFilter : function(req, file, callback) { //file filter
                    if (['xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length-1]) === -1) {
                        return callback(new Error('Wrong extension type'));
                    }
                    callback(null, true);
                }
            }).single('file');

app.post('/upload', function(req, res) {
    upload(req,res,function(err){
        if(err){
                res.json({error_code:1,err_desc:err});
                return;
        }
            res.json({error_code:0,err_desc:null});
            const newdata = excelToJson({
                sourceFile: req.file.path,
                columnToKey: {
                    '*': '{{columnHeader}}'
                }
            });
            MongoClient.connect(url, function(err, db) {
                if (err) throw err;
                var dbo = db.db("mydb");
            
                // This is to output data, set query with "var query = {}"" to output all
                // https://www.w3schools.com/nodejs/nodejs_mongodb_query.asp
                var query = {};
                dbo.collection("customers").find(query).toArray(function(err, result) {
                    if (err) throw err;
                    console.log(result);
                    db.close();
                });
            
                // Update database with newdata
                for (let i = 1, l = newdata.Sheet1.length; i < l; i++) {
                    if (typeof newdata.Sheet1[i]["Stock"] != "number") {
                        console.log("[ERROR] Stok kode SKU " + newdata.Sheet1[i]["Kode SKU"] + " bukan angka");
                    } else {
                        var query = { "Kode SKU" : newdata.Sheet1[i]["Kode SKU"]};
                        var newvalue = {$set : { Stock : newdata.Sheet1[i]["Stock"]}};
                        dbo.collection("customers").updateOne(query,newvalue, function(err, res){
                            if (err) throw err;
                            if (res.result.n != 0){
                                console.log("[SUCCESS] Stok kode SKU " + newdata.Sheet1[i]["Kode SKU"] + " berhasil di-update");
                            }
                            db.close();
                        });
                    }
                }
            
                // This is to output data, set query with "var query = {}"" to output all
                // https://www.w3schools.com/nodejs/nodejs_mongodb_query.asp
                var query = {};
                dbo.collection("customers").find(query).toArray(function(err, result) {
                    if (err) throw err;
                    console.log(result);
                    db.close();
                });
            
                // This it to delete data, set query with "var query ={}" to delete all
                // https://www.w3schools.com/nodejs/nodejs_mongodb_remove.asp
                // var query = { Stock : 'Stock'}
                // dbo.collection("customers").remove(query, function(err, obj) {
                //     if (err) throw err;
                //     console.log(obj.result.n + " document(s) deleted");
                //     db.close();
                // });
            });
            
    });
});

app.get('/',function(req,res){
res.sendFile(__dirname + "/index.html");
});

app.listen('3000', function(){
    console.log('running on 3000...');
});