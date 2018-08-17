var express = require("express");
var app = express();
var bodyParser = require("body-parser");
var multer = require("multer");
var xlstojson = require("xls-to-json-lc");
var _ = require("lodash");

app.set("view engine", "ejs");
app.use(bodyParser.json());

var storage = multer.diskStorage({
  //multers disk storage settings
  destination: function(req, file, cb) {
    cb(null, "./uploads/");
  },
  filename: function(req, file, cb) {
    var datetimestamp = Date.now();
    cb(
      null,
      file.fieldname +
        "-" +
        datetimestamp +
        "." +
        file.originalname.split(".")[file.originalname.split(".").length - 1]
    );
  }
});

var upload = multer({
  //multer settings
  storage: storage,
  fileFilter: function(req, file, callback) {
    //file filter
    if (
      ["xls", "xlsx"].indexOf(
        file.originalname.split(".")[file.originalname.split(".").length - 1]
      ) === -1
    ) {
      return callback(new Error("Wrong extension type"));
    }
    callback(null, true);
  }
}).single("file");

/** API path that will upload the files */
app.post("/upload", function(req, res) {
  var exceltojson;
  upload(req, res, function(err) {
    if (err) {
      res.json({ error_code: 1, err_desc: err });
      return;
    }
    /** Multer gives us file info in req.file object */
    if (!req.file) {
      res.json({ error_code: 1, err_desc: "No file passed" });
      return;
    }
    /** Check the extension of the incoming file and
     *  use the appropriate module
     */
    if (
      req.file.originalname.split(".")[
        req.file.originalname.split(".").length - 1
      ] === "xls"
    ) {
      exceltojson = xlstojson;
    }
    console.log(req.file.path);
    try {
      exceltojson(
        {
          input: req.file.path,
          output: null, //since we don't need output.json
          lowerCaseHeaders: true
        },
        function(err, result) {
          if (err) {
            return res.json({ error_code: 1, err_desc: err, data: null });
          }
          var sp = [];
          _.forEach(result, function(value) {
            // console.log(value);
            if (value["tình trạng đơn hàng"] === "Chờ giao hàng") {
              const totalSP = value["thông tin sản phẩm"].split("\n");
              _.forEach(totalSP, function(value2) {
                sp.push(value2);
              });
            }
          });
          //   var arrSp1 = valsp1.split(";");
          var sp1 = [];
          _.forEach(sp, function(valsp1) {
            var arrSp1 = valsp1.split(";");

            name = "";
            sl = 0;
            phanLoai = "";
            _.forEach(arrSp1, function(valsp2) {
              if (valsp2.includes("] Tên phân loại hàng:")) {
                name = valsp2.substring(23);
              }
              if (valsp2.startsWith(" Tên phân loại hàng:")) {
                phanLoai = valsp2.replace(" Tên phân loại hàng:", "");
              }
              if (valsp2.startsWith(" Số lượng:")) {
                sl = parseInt(valsp2.replace(" Số lượng: ", ""));
              }
            });
            sp1.push({ name, phanLoai, sl });
          });

          var newSp = [];
          _.forEach(sp1, value => {
            newSp.push({ name: value.name, phanLoai: value.phanLoai });
          });
          newSp = _.uniqWith(newSp, _.isEqual);

          newSp = _.orderBy(newSp, ["name"], ["asc"]);

          var newSp1 = [];
          for (var i = 0; i < newSp.length; i++) {
            var sl = 0;
            _.forEach(sp1, value => {
              if (
                newSp[i].name === value.name &&
                newSp[i].phanLoai === value.phanLoai
              ) {
                sl += value.sl;
              }
            });

            newSp1.push({
              name: newSp[i].name,
              phanLoai: newSp[i].phanLoai,
              sl
            });
          }
          //   res.json({ newSp1 });
          console.log(newSp1);
          res.render("index", { newSp1 });
        }
      );
    } catch (e) {
      res.json({ error_code: 1, err_desc: "Corupted excel file" });
    }
  });
});

app.get("/", function(req, res) {
  res.sendFile(__dirname + "/index.html");
});

app.listen("3000", function() {
  console.log("running on 3000...");
});
