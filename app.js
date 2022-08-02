const express = require("express");
const mongoose = require("mongoose");
const bodyParser = require("body-parser");

const Users = require("./src/users/user.model");
const ExcelJS = require("exceljs");
const moment = require("moment");

require("dotenv").config();
const PORT = 5000;

const authRoutes = require("./routes/users");

mongoose
  .connect(process.env.MONGO_URI, {
    useNewUrlParser: true,
    useUnifiedTopology: true,
  })
  .then(() => {
    console.log("Database connection Success.");
  })
  .catch((err) => {
    console.error("Mongo Connection Error", err);
  });

const app = express();

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/ping", (req, res) => {
  return res.send({
    error: false,
    message: "Server is healthy",
  });
});

app.use("/users", authRoutes);

app.listen(PORT, () => {
  console.log("Server started listening on PORT : " + PORT);
});

app.get("/sheet", async (req, res, next) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");
  worksheet.columns = [
    { header: "Name", key: "name", width: 30 }, 
    { header: "Email", key: "email", width: 30 },
    { header: "Password", key: "password", width: 30 },
    { header: "Referrer", key: "referrer", width: 30 },
    { header: "Created At", key: "createdAt", width: 30 },
  ];
  try{
  const Users = await Users.find({});
  Users.forEach((Users) => {
    worksheet.addRow({
      name: Users.name,
      email: Users.email,
      password: Users.password,
      referrer: Users.referrer,
      createdAt: moment(Users.createdAt).format("DD-MM-YYYY"),
    });
  }
  );
  workbook.xlsx.writeFile("users.xlsx").then(() => {
    res.send("File created");
    console.log("File written successfully");
  }
  );
  }
  catch(err){
    console.log(err);
    res.status(500).send(err);
  }
}
);
