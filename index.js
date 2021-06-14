const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const app = require('express')();
const json2xls = require('json2xls');
const multer = require("multer");
const path = require('path');


app.use(express.static(__dirname));


const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, './uploads') //Destination folder
  },
  filename: function (req, file, cb) {
    cb(null, 'price.xlsx') //File name after saving
  }
})

const upload = multer({storage: storage})


const convertFile = (res) => {
  const fileNameUploads = "price.xlsx";
  const fileUploads = __dirname + "/uploads/" + fileNameUploads;
  const workbook = XLSX.readFile(fileUploads);
  const sheet_name_list = workbook.SheetNames;
  const xlData = XLSX.utils.sheet_to_json(
    workbook.Sheets[sheet_name_list[0]]
  );

  for (let value of xlData) {
    const VO = 'VO';
    const RT = 'RT';
    value['Код товара'] = value['Код товара'].replace(VO, '');
    value['Код товара'] = value['Код товара'].replace(RT, '');
  }

  const filename = 'priceVOLVO.xlsx';
  const convert = () => {
    const xls = json2xls(xlData);
    fs.writeFileSync(filename, xls, 'binary', () => {
      console.log(filename + ' file is saved!');
    });
  };
  convert();

  const file = `${__dirname}/priceVOLVO.xlsx`;
  res.download(file);
}


app.get('/', function (req, res) {
  res.sendFile(__dirname + '/index.html');
});


app.post('/', upload.single('data'), (req, res) => {
  const data = req.file;
  if (!data)
    res.send("Ошибка при загрузке файла");
  else
  convertFile(res)
  // res.redirect('/');
});

const port = process.env.PORT || 5000;
app.listen(port, () => {
  console.log(`app is running on ${port}`);
});
