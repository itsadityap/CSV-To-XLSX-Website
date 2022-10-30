const express = require("express");
const fileUpload = require("express-fileupload");
const path = require("path");
const filesPayloadExists = require('./middleware/filesPayloadExists');
const fs = require('fs');
const csvtojson = require('csvtojson');
const XLSX = require('xlsx');
const fileExtLimiter = require("./middleware/fileExtLimiter");

const PORT = process.env.PORT || 3000;

const app = express();

//GET ROUTES
app.get("/", (req,res) => 
{
    res.sendFile(path.join(__dirname, "index.html"));
})

//FUNCTIONS
const ReadCSV = async () => {

    const csvfilepath = "./files/sales_data_sample.csv"
    csvtojson()
    .fromFile(csvfilepath)
    .then((json) => {
        fs.writeFileSync('./files/op.json', JSON.stringify(json));
    })
    .catch((err) => {console.log(err)})
}

const JSONToSpecificJSON = async () => {

    let proline= []
    let terr = []
    let sal = []
    let netSal = []
    let totSalAmt = []

    let ourData = []

    let rawdata = fs.readFileSync('./files/op.json');
    let rawdataJSON = JSON.parse(rawdata);

    for(let i=0;i<Object.keys(rawdataJSON).length;i++)
    {    
        if(rawdataJSON[i].YEAR_ID == 2004)
        {
            proline.push(rawdataJSON[i].PRODUCTLINE)
            terr.push(rawdataJSON[i].TERRITORY)
            sal.push(rawdataJSON[i].QUANTITYORDERED)
            netSal.push(rawdataJSON[i].SALES-rawdataJSON[i].ORDERLINENUMBER)
            totSalAmt.push(rawdataJSON[i].MSRP*(rawdataJSON[i].SALES-rawdataJSON[i].ORDERLINENUMBER))
        }
    }

    for(let i=0;i<proline.length;i++)
    {    
        ourData.push({
            "PRODUCT_LINE":proline[i],
            "TERRITORY":terr[i],
            "QUANTITYORDERED":sal[i],
            "NET_SALES":netSal[i],
            "TOT_AMT_SALES":totSalAmt[i],
        })
    }

    fs.writeFileSync('./files/final.json', JSON.stringify(ourData));
}

const JSONToExcel = () => {

    let finaldata = fs.readFileSync('./files/final.json');
    let finaldataJSON = JSON.parse(finaldata);

    const workSheet = XLSX.utils.json_to_sheet(finaldataJSON);
    const workBook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workBook, workSheet, "DataExcel")
    
    // Generate buffer
    XLSX.write(workBook, { bookType: 'xlsx', type: "buffer" })

    // Binary string
    XLSX.write(workBook, { bookType: "xlsx", type: "binary" })

    XLSX.writeFile(workBook, "Data.xlsx")

}

// const DeleteFileOP = async () => 
// {
//     try 
//     {
//         fs.unlinkSync("./files/op.json")
//         console.log("JSON File is Deleted");
//     }
//     catch(err)
//     {
//         console.log("Cannot Delete File Doesn't Exist");
//     }
// }

const DeleteFileExcel = async () => 
{
    try 
    {
        fs.unlinkSync("Data.xlsx")
        console.log("File is Deleted");
    }
    catch(err)
    {
        console.log("Cannot Delete File Doesn't Exist");
    }
}

const DeleteFileCSV = async () => 
{
    try 
    {
        fs.unlinkSync("./files/sales_data_sample.csv")
        console.log("CSV File Deleted");
    }
    catch(err)
    {
        console.log("Cannot Delete File Doesn't Exist");
    }
}

//POST ROUTES
app.post('/upload',
        fileUpload({ createParentPath: true }),
        fileExtLimiter(['.csv']),
        filesPayloadExists,
        (req, res) => {
            const files = req.files
            console.log(files)

            Object.keys(files).forEach(key => {
                const filepath = path.join(__dirname, 'files', files[key].name)
                files[key].mv(filepath, (err) => {
                    if (err) return res.status(500).json({ status: "error", message: err })
                })
            })

            ReadCSV()

            JSONToSpecificJSON()

            JSONToExcel()

            return res.json({ status: 'success', message: Object.keys(files).toString() })
            //res.redirect('/download')
    }
)

app.post("/donwloads", (req,res) => {

        const filePath = 'Data.xlsx' 

        res.download(
            filePath, 
            "Data-Excel.xlsx",

            (err) => {
                if (err) {
                    res.send({
                        error : err,
                        msg   : "Problem downloading the file"
                    })
                }
        });
        
        setTimeout(DeleteFileExcel,8000)
        //setTimeout(DeleteFileOP,9000)
        setTimeout(DeleteFileCSV,10000)
})

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));