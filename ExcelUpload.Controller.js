let express = require("express");
const app = express();
const router = express.Router();
let upload = require("express-fileupload");
let importExcel = require("convert-excel-to-json");
const mongoose = require("mongoose");

const jurisdictionScehma = require("../models/Jurisdiction.Model");
const legalStatusScheam = require("../models/LegalStatus.Model");
const companySchema = require("../models/Company.Model");

const fs = require("fs");
if (!fs.existsSync('./excel')) {
  fs.mkdir("./excel", function (err) {
    if (err) {
      console.log(err);
    } else {
      console.log("New directory successfully created.");
    }
  });
}

router.use(upload());
app.use(express.json());

router.post("/api/upload-excel", function (req, res) {
  let file = req.files.filename;
  let filename = file.name;

  let data = [];
  let JurisdictionList = [];
  let LegalStatusList = [];

  file.mv("./excel/" + filename, (err) => {
    if (err) {
      res.status(400).json({
        message: err.message,
      });
    } else {
      let result = importExcel({
        sourceFile: "./excel/" + filename,
        header: { rows: 1 },
        columnToKey: {
          A: "Company_Name",
          B: "Company_Email",
          C: "Jurisdiction",
          D: "LegalStatus",
          E: "LAST_FEE",
          F: "FEE",
          G: "NTN",
          H: "STN",
          I: "Branch_Name",
          J: "Fax",
          K: "BulkEmail",
          L: "Contact_Personal_Name",
          M: "Contact_Personal_Email",
          N: "Contact_Personal_Phone",
          O: "Contact_Personal_FaxNo",
          P: "Contact_Personal_Cell_1",
          Q: "Contact_Personal_Cell_2",
          R: "Contact_Personal_Address_1",
          S: "Contact_Personal_Address_2",
          T: "ArrayOfRetainer",
          U: "ArrayOfNoNRetainer",
          V: "IsActive",
        },
        sheets: ["Sheet1"],
      });

      for (let i = 0; result.Sheet1.length > i; i++) {
        jurisdictionScehma
          .find({ Jurisdiction: { $regex: new RegExp("^" + result.Sheet1[i].Jurisdiction, "i") }})
          .then((jurisdiction) => {
            jurisdiction.map((item) => {
              if (result.Sheet1[i].Jurisdiction.toString().toUpperCase() == item.Jurisdiction.toString().toUpperCase()) {
                JurisdictionList.push({
                  Jurisdiction: new mongoose.Types.ObjectId(item._id),
                });
              }
              //  legal status colunm id fetched by given legal status name in excel sheet
              legalStatusScheam
                .find({ LegalStatus: { $regex: new RegExp("^" + result.Sheet1[i].LegalStatus, "i")}})
                .then((legal_status) => {
                  legal_status.map((item, idx) => {
                    if (result.Sheet1[i].LegalStatus.toString().toUpperCase() == item.LegalStatus.toString().toUpperCase()) {
                      LegalStatusList.push({
                        LegalStatus: new mongoose.Types.ObjectId(item._id),
                      });
                    }
                    new Promise((resolve, reject) => {
                      data.push({
                        Company_Name: result.Sheet1[i].Company_Name ?? "",
                        Company_Email: result.Sheet1[i].Company_Email,
                        Jurisdiction: JurisdictionList[idx]["Jurisdiction"],
                        LegalStatus: LegalStatusList[idx]["LegalStatus"],
                        LAST_FEE: result.Sheet1[i].LAST_FEE,
                        FEE: result.Sheet1[i].FEE,
                        NTN: result.Sheet1[i].NTN,
                        STN: result.Sheet1[i].STN,
                        Branch_Name: result.Sheet1[i].Branch_Name,
                        Fax: result.Sheet1[i].Fax,
                        BulkEmail: [],
                        ArrayOfRetainer: [],
                        ArrayOfNoNRetainer: [],
                        Contact_Personal_Name: "",
                        Contact_Personal_Email: "",
                        Contact_Personal_Phone: "",
                        Contact_Personal_FaxNo: "",
                        Contact_Personal_Cell_1: "",
                        Contact_Personal_Cell_2: "",
                        Contact_Personal_Address_1: "",
                        Contact_Personal_Address_2: "",
                        IsRetainerActvie: false,
                        IsNonRetainerActvie: false,
                        IsActive: true,
                      });
                      if (data.length) {
                        resolve(data);
                      } else {
                        reject();
                      }
                    }).then((array) => {
                      for (let index = 0; index < array.length; index++) {
                        if (index == array.length - 1) {
                          if (array.length == result.Sheet1.length) {
                            companySchema
                              .insertMany(array)
                              .then((data) => {
                                return res.status(200).json({
                                  status: "success",
                                  message: "Company Added Successfully!",
                                  data: data,
                                });
                              })
                              .catch((err) => {
                                res.status(400).json({
                                  status: "failed",
                                  message: "Something Went Wrong Kindly check your excel format please ",
                                });
                              });
                          }
                        }
                      }
                    });
                  });
                });
            });
          });
      }
    }
  });
});

module.exports = router;
