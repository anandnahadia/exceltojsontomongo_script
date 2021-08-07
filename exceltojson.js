//You need to have mongodb installed in your system
//Run npm start/node index.js to run the script


const excelToJson = require('convert-excel-to-json');
let MongoClient = require('mongodb').MongoClient;
let url = "mongodb://localhost:27017/";

const result = excelToJson({
    sourceFile: 'Test Hospital Data.xlsx'
});

//Json which will be ingested in database

let hospital_data = []

for(let i=1;i<result['Multi Speciality Marketing'].length;i++){

    hospital_data[i-1] = {}
    hospital_data[i-1]['Hospital Name']= result['Multi Speciality Marketing'][i].A
    hospital_data[i-1]['Hospital Details']= result['Multi Speciality Marketing'][i].N
    hospital_data[i-1]['Mobile Number']= result['Multi Speciality Marketing'][i].B
    hospital_data[i-1]['Address']= result['Multi Speciality Marketing'][i].C

    hospital_data[i-1].Doctors = {
        'Doctor Name' : result['Multi Speciality Marketing'][i].D,
        'Qualification' : (result['Multi Speciality Marketing'][i].E).split(","),
        'Specialization' : result['Multi Speciality Marketing'][i].F,
        'Speciality (According to admin panel)' : result['Multi Speciality Marketing'][i].G,
        'Experience' : result['Multi Speciality Marketing'][i].H,
        'Consultation Fees' : result['Multi Speciality Marketing'][i].I,
        'Days' : result['Multi Speciality Marketing'][i].J,
        'Timings' : (result['Multi Speciality Marketing'][i].K).split('&'),
        'Designation' : result['Multi Speciality Marketing'][i].L,
        'Doctor Details' : result['Multi Speciality Marketing'][i].M
    }

}
//Insert Json into Mongodb Database Collection

MongoClient.connect(url, { useNewUrlParser: true }, (err, db) => {
    if (err) throw err;
    
    var dbo = db.db("hospital_information_database");
    
    dbo.collection("hospital_information").insertMany(hospital_data, (err, res) => {
      if (err) throw err;
      
      console.log("Number of Rows inserted: " + res.insertedCount);
      
      db.close();
    });
});

// //Fetching Data From Mongodb Database Collection

MongoClient.connect(url, function(err, db) {
    if (err) throw err;
    var dbo = db.db("hospital_information_database");
    dbo.collection("hospital_information").find({}).toArray(function(err, result) {
      if (err) throw err;
      console.log(result);
      db.close();
    });
});

