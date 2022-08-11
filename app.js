const XLSX = require('xlsx');
const path = require('path');
const POST = 3000;
const express = require('express');
const { read, readFile } = require('fs');
const app = express();
app.set('view engine', 'ejs')
const staticWebsite = path.join(__dirname , './views')

app.use(express.static(staticWebsite))


bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: false }));




// functions

// function that create a new sheet

const exportUsersToExcel = (UserList, workSheetName, filePath) => {

    const workBook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(UserList)
    XLSX.utils.book_append_sheet(workBook, worksheet ,workSheetName , true)
    XLSX.writeFile(workBook, path.resolve(filePath));
    return true;

};


// function that append data to a existing sheet


// const addToExistingSheet = (UserList, filePath, dogsName) => {
//     const workBook = XLSX.readFile('users.xlsx');
//     const worksheet = workBook.Sheets[dogsName];
//     XLSX.utils.sheet_add_json(worksheet, UserList, {skipHeader:true , origin:-1});
//     XLSX.writeFile(workBook, path.resolve(filePath));

// }





let UserList = [];


const filePath = './users.xlsx';




// requests

app.post('' , (req , res) => {
    res.render('index');
    UserList = [{
        'dog' : `${req.body.dog}`,
        'handler' : `${req.body.handler}`,
        'day' : `${req.body.day}`,
        'weather' : `${req.body.weather}`,
        'food' : `${req.body.food}`,
        'hour' : `${req.body.hour}`,
        'sensetivity' : `${req.body.sensetivity}`,
        'specificity' : `${req.body.specificity}`,
        'ppv' : `${req.body.ppv}`,
        'npv' : `${req.body.npv}`
    
    }];
    const workSheetName = `${req.body.dog}`;



    // console.log(UserList);


        exportUsersToExcel( UserList , workSheetName , filePath);

});




app.get('' , (req , res) => {
    res.render('index')}
);


app.listen(POST , () => console.log('the app is listening to port 3000!!'));
const XLSX = require('xlsx');
const path = require('path');
const POST = 3000;
const express = require('express');
const { read, readFile } = require('fs');
const app = express();
app.set('view engine', 'ejs')
const staticWebsite = path.join(__dirname , './views')

app.use(express.static(staticWebsite))


bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: false }));




// functions

// function that create a new sheet

const exportUsersToExcel = (UserList, workSheetName, filePath) => {

    const workBook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(UserList)
    XLSX.utils.book_append_sheet(workBook, worksheet ,workSheetName , true)
    XLSX.writeFile(workBook, path.resolve(filePath));
    return true;

};


// function that append data to a existing sheet


// const addToExistingSheet = (UserList, filePath, dogsName) => {
//     const workBook = XLSX.readFile('users.xlsx');
//     const worksheet = workBook.Sheets[dogsName];
//     XLSX.utils.sheet_add_json(worksheet, UserList, {skipHeader:true , origin:-1});
//     XLSX.writeFile(workBook, path.resolve(filePath));

// }





let UserList = [];


const filePath = './users.xlsx';




// requests

app.post('' , (req , res) => {
    res.render('index');
    UserList = [{
        'dog' : `${req.body.dog}`,
        'handler' : `${req.body.handler}`,
        'day' : `${req.body.day}`,
        'weather' : `${req.body.weather}`,
        'food' : `${req.body.food}`,
        'hour' : `${req.body.hour}`,
        'sensetivity' : `${req.body.sensetivity}`,
        'specificity' : `${req.body.specificity}`,
        'ppv' : `${req.body.ppv}`,
        'npv' : `${req.body.npv}`
    
    }];
    const workSheetName = `${req.body.dog}`;



    // console.log(UserList);


        exportUsersToExcel( UserList , workSheetName , filePath);

});




app.get('' , (req , res) => {
    res.render('index')}
);


app.listen(POST , () => console.log('the app is listening to port 3000!!'));
const XLSX = require('xlsx');
const path = require('path');
const POST = 3000;
const express = require('express');
const { read, readFile } = require('fs');
const app = express();
app.set('view engine', 'ejs')
const staticWebsite = path.join(__dirname , './views')

app.use(express.static(staticWebsite))


bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: false }));




// functions

// function that create a new sheet

const exportUsersToExcel = (UserList, workSheetName, filePath) => {

    const workBook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(UserList)
    XLSX.utils.book_append_sheet(workBook, worksheet ,workSheetName , true)
    XLSX.writeFile(workBook, path.resolve(filePath));
    return true;

};


// function that append data to a existing sheet


// const addToExistingSheet = (UserList, filePath, dogsName) => {
//     const workBook = XLSX.readFile('users.xlsx');
//     const worksheet = workBook.Sheets[dogsName];
//     XLSX.utils.sheet_add_json(worksheet, UserList, {skipHeader:true , origin:-1});
//     XLSX.writeFile(workBook, path.resolve(filePath));

// }





let UserList = [];


const filePath = './users.xlsx';




// requests

app.post('' , (req , res) => {
    res.render('index');
    UserList = [{
        'dog' : `${req.body.dog}`,
        'handler' : `${req.body.handler}`,
        'day' : `${req.body.day}`,
        'weather' : `${req.body.weather}`,
        'food' : `${req.body.food}`,
        'hour' : `${req.body.hour}`,
        'sensetivity' : `${req.body.sensetivity}`,
        'specificity' : `${req.body.specificity}`,
        'ppv' : `${req.body.ppv}`,
        'npv' : `${req.body.npv}`
    
    }];
    const workSheetName = `${req.body.dog}`;



    // console.log(UserList);


        exportUsersToExcel( UserList , workSheetName , filePath);

});




app.get('' , (req , res) => {
    res.render('index')}
);


app.listen(POST , () => console.log('the app is listening to port 3000!!'));
const XLSX = require('xlsx');
const path = require('path');
const POST = 3000;
const express = require('express');
const { read, readFile } = require('fs');
const app = express();
app.set('view engine', 'ejs')
const staticWebsite = path.join(__dirname , './views')

app.use(express.static(staticWebsite))


bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: false }));




// functions

// function that create a new sheet

const exportUsersToExcel = (UserList, workSheetName, filePath) => {

    const workBook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(UserList)
    XLSX.utils.book_append_sheet(workBook, worksheet ,workSheetName , true)
    XLSX.writeFile(workBook, path.resolve(filePath));
    return true;

};


// function that append data to a existing sheet


// const addToExistingSheet = (UserList, filePath, dogsName) => {
//     const workBook = XLSX.readFile('users.xlsx');
//     const worksheet = workBook.Sheets[dogsName];
//     XLSX.utils.sheet_add_json(worksheet, UserList, {skipHeader:true , origin:-1});
//     XLSX.writeFile(workBook, path.resolve(filePath));

// }





let UserList = [];


const filePath = './users.xlsx';




// requests

app.post('' , (req , res) => {
    res.render('index');
    UserList = [{
        'dog' : `${req.body.dog}`,
        'handler' : `${req.body.handler}`,
        'day' : `${req.body.day}`,
        'weather' : `${req.body.weather}`,
        'food' : `${req.body.food}`,
        'hour' : `${req.body.hour}`,
        'sensetivity' : `${req.body.sensetivity}`,
        'specificity' : `${req.body.specificity}`,
        'ppv' : `${req.body.ppv}`,
        'npv' : `${req.body.npv}`
    
    }];
    const workSheetName = `${req.body.dog}`;



    // console.log(UserList);


        exportUsersToExcel( UserList , workSheetName , filePath);

});




app.get('' , (req , res) => {
    res.render('index')}
);


app.listen(POST , () => console.log('the app is listening to port 3000!!'));
