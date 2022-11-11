const express = require("express");
const PORT = process.env.PORT || 8031;
const app = express();
const pug = require("pug");
const path = require("path");
const xl = require("excel4node");
const wb = new xl.Workbook();

app.set("view engine", "pug");
app.set("views", path.join(__dirname, "./views"));
app.use(express.static("./public"));

const MongoClient = require("mongodb").MongoClient;
const db_url =
  "mongodb+srv://sandeep:aZMYo0C6G3KxYqMA@cluster-details.ffmkvzn.mongodb.net/?retryWrites=true&w=majority";

app.get("/", (req, res) => {
  res.render("home.pug");
});

app.get("/event1", (req, res) => {
  MongoClient.connect(db_url, async (err, db) => {
    if (err) throw err;
    const dbo = db.db("event-details");
    let cursor1 = await dbo.collection("event1").find({});
    let values1 = await cursor1.toArray();
    let details = [];
    values1.forEach((element) => {
      details.push({
        team_name: element["team_name"],
        name1: element["name1"],
        rollno: element["rollno"],
        email: element["email"],
        mobileno: element["mobileno"],
        title: element["title"],
        abstract: element["abstract"],
        name2: element["name2"],
        rollno2: element["rollno2"],
        name3: element["name3"],
        rollno3: element["rollno3"],
      });
    });

    const headingColumnNames = [
      "team_name",
      "name1",
      "rollno",
      "email",
      "mobileno",
      "title",
      "abstract",
      "name2",
      "rollno2",
      "name3",
      "rollno3",
    ];
    //Write Column Title in Excel file
    const ws1 = wb.addWorksheet("Worksheet 1");

    let headingColumnIndex = 1;
    headingColumnNames.forEach((heading) => {
      ws1.cell(1, headingColumnIndex++).string(heading);
    });
    //Write Data in Excel file
    let rowIndex = 2;
    details.forEach((record) => {
      let columnIndex = 1;
      Object.keys(record).forEach((columnName) => {
        ws1.cell(rowIndex, columnIndex++).string(record[columnName]);
      });
      rowIndex++;
    });
    wb.write("event1.xlsx", () => {
      const file = `${__dirname}/event1.xlsx`;
      res.download(file);
    });
    db.close();
  });
});

app.get("/event2", (req, res) => {
  MongoClient.connect(db_url, async (err, db) => {
    if (err) throw err;
    const dbo = db.db("event-details");
    let cursor1 = await dbo.collection("event2").find({});
    let values1 = await cursor1.toArray();
    let details = [];
    values1.forEach((element) => {
      details.push({
        team_name: element["team_name"],
        name1: element["name1"],
        rollno: element["rollno"],
        email: element["email"],
        mobileno: element["mobileno"],
        title: element["title"],
        abstract: element["abstract"],
        name2: element["name2"],
        rollno2: element["rollno2"],
        name3: element["name3"],
        rollno3: element["rollno3"],
      });
    });

    const headingColumnNames = [
      "team_name",
      "name1",
      "rollno",
      "email",
      "mobileno",
      "title",
      "abstract",
      "name2",
      "rollno2",
      "name3",
      "rollno3",
    ];
    //Write Column Title in Excel file
    const ws2 = wb.addWorksheet("Worksheet 2");

    let headingColumnIndex = 1;
    headingColumnNames.forEach((heading) => {
      ws2.cell(1, headingColumnIndex++).string(heading);
    });
    //Write Data in Excel file
    let rowIndex = 2;
    details.forEach((record) => {
      let columnIndex = 1;
      Object.keys(record).forEach((columnName) => {
        ws2.cell(rowIndex, columnIndex++).string(record[columnName]);
      });
      rowIndex++;
    });
    wb.write("event2.xlsx", () => {
      const file = `${__dirname}/event2.xlsx`;
      res.download(file);
    });
    db.close();
  });
});

app.get("/event3", (req, res) => {
  MongoClient.connect(db_url, async (err, db) => {
    if (err) throw err;
    const dbo = db.db("event-details");
    let cursor1 = await dbo.collection("event3").find({});
    let values1 = await cursor1.toArray();
    let details = [];
    values1.forEach((element) => {
      details.push({
        name1: element["name"],
        rollno: element["rollno"],
        email: element["email"],
        mobileno: element["mobileno"],
      });
    });

    const headingColumnNames = ["name", "rollno", "email", "mobileno"];
    //Write Column Title in Excel file
    const ws3 = wb.addWorksheet("Worksheet 3");

    let headingColumnIndex = 1;
    headingColumnNames.forEach((heading) => {
      ws3.cell(1, headingColumnIndex++).string(heading);
    });
    //Write Data in Excel file
    let rowIndex = 2;
    details.forEach((record) => {
      let columnIndex = 1;
      Object.keys(record).forEach((columnName) => {
        ws3.cell(rowIndex, columnIndex++).string(record[columnName]);
      });
      rowIndex++;
    });
    wb.write("event3.xlsx", () => {
      const file = `${__dirname}/event3.xlsx`;
      res.download(file);
    });

    db.close();
  });
});

app.get("/event4", (req, res) => {
  MongoClient.connect(db_url, async (err, db) => {
    if (err) throw err;
    const dbo = db.db("event-details");
    let cursor1 = await dbo.collection("event4").find({});
    let values1 = await cursor1.toArray();
    let details = [];
    values1.forEach((element) => {
      details.push({
        name1: element["name"],
        rollno: element["rollno"],
        email: element["email"],
        mobileno: element["mobileno"],
      });
    });

    const headingColumnNames = ["name", "rollno", "email", "mobileno"];
    //Write Column Title in Excel file
    const ws4 = wb.addWorksheet("Worksheet 4");

    let headingColumnIndex = 1;
    headingColumnNames.forEach((heading) => {
      ws4.cell(1, headingColumnIndex++).string(heading);
    });
    //Write Data in Excel file
    let rowIndex = 2;
    details.forEach((record) => {
      let columnIndex = 1;
      Object.keys(record).forEach((columnName) => {
        ws4.cell(rowIndex, columnIndex++).string(record[columnName]);
      });
      rowIndex++;
    });
    wb.write("event4.xlsx", () => {
      const file = `${__dirname}/event4.xlsx`;
      res.download(file);
    });

    db.close();
  });
});

app.get("/event5", (req, res) => {
  MongoClient.connect(db_url, async (err, db) => {
    if (err) throw err;
    const dbo = db.db("event-details");
    let cursor1 = await dbo.collection("event5").find({});
    let values1 = await cursor1.toArray();
    let details = [];
    values1.forEach((element) => {
      details.push({
        team_name: element["team_name"],
        name1: element["name1"],
        rollno: element["rollno"],
        email: element["email"],
        mobileno: element["mobileno"],
        name2: element["name2"],
        rollno2: element["rollno2"],
        name3: element["name3"],
        rollno3: element["rollno3"],
      });
    });

    const headingColumnNames = [
      "team_name",
      "name1",
      "rollno",
      "email",
      "mobileno",
      "name2",
      "rollno2",
      "name3",
      "rollno3",
    ];
    //Write Column Title in Excel file
    const ws5 = wb.addWorksheet("Worksheet 5");

    let headingColumnIndex = 1;
    headingColumnNames.forEach((heading) => {
      ws5.cell(1, headingColumnIndex++).string(heading);
    });
    //Write Data in Excel file
    let rowIndex = 2;
    details.forEach((record) => {
      let columnIndex = 1;
      Object.keys(record).forEach((columnName) => {
        ws5.cell(rowIndex, columnIndex++).string(record[columnName]);
      });
      rowIndex++;
    });
    wb.write("event5.xlsx", () => {
      const file = `${__dirname}/event5.xlsx`;
      res.download(file);
    });

    db.close();
  });
});

app.get("/event6", (req, res) => {
  MongoClient.connect(db_url, async (err, db) => {
    if (err) throw err;
    const dbo = db.db("event-details");
    let cursor1 = await dbo.collection("event6").find({});
    let values1 = await cursor1.toArray();
    let details = [];
    values1.forEach((element) => {
      details.push({
        team_name: element["team_name"],
        name1: element["name1"],
        rollno: element["rollno"],
        email: element["email"],
        mobileno: element["mobileno"],
        name2: element["name2"],
        rollno2: element["rollno2"],
        name3: element["name3"],
        rollno3: element["rollno3"],
      });
    });

    const headingColumnNames = [
      "team_name",
      "name1",
      "rollno",
      "email",
      "mobileno",
      "name2",
      "rollno2",
      "name3",
      "rollno3",
    ];
    //Write Column Title in Excel file
    const ws6 = wb.addWorksheet("Worksheet 6");

    let headingColumnIndex = 1;
    headingColumnNames.forEach((heading) => {
      ws6.cell(1, headingColumnIndex++).string(heading);
    });
    //Write Data in Excel file
    let rowIndex = 2;
    details.forEach((record) => {
      let columnIndex = 1;
      Object.keys(record).forEach((columnName) => {
        ws6.cell(rowIndex, columnIndex++).string(record[columnName]);
      });
      rowIndex++;
    });
    wb.write("event6.xlsx", () => {
      const file = `${__dirname}/event6.xlsx`;
      res.download(file);
    });

    db.close();
  });
});

app.listen(PORT, () => {
  console.log(`Server listening on ${PORT}`);
});
