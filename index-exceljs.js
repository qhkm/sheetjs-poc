const fastify = require("fastify")({ logger: true });
const ExcelJS = require("exceljs");

const fs = require("fs");
const axios = require("axios");

// Declare a route
fastify.get("/", async (request, reply) => {
  // var fileContents = Buffer.from(fileData, "base64");
  const data = [
    {
      name: "yum",
      id: 1,
    },
  ];

  // CREATE WORKBOOK
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Me";
  workbook.lastModifiedBy = "Her";
  workbook.created = new Date(1985, 8, 30);
  workbook.modified = new Date();
  workbook.lastPrinted = new Date(2016, 9, 27);

  // add workbook views
  workbook.views = [
    {
      x: 0,
      y: 0,
      width: 10000,
      height: 20000,
      firstSheet: 0,
      activeTab: 1,
      visibility: "visible",
    },
  ];

  // ADD WORKSHEET
  const worksheet = workbook.addWorksheet("My Sheet");
  worksheet.columns = [
    { header: "Id", key: "id" },
    { header: "Name", key: "name" },
    { header: "Age", key: "age" },
  ];

  // ADD ROW: method 1
  worksheet.addRow({ id: 1, name: "John Doe", age: new Date(1970, 1, 1) });
  worksheet.addRow({ id: 2, name: "Jane Doe", age: new Date(1965, 1, 7) });

  // ADD ROW: method 2
  const rows = [[3, "Alex", "44"], { id: 4, name: "Margaret", age: 32 }];
  worksheet.addRows(rows);

  // ADD PAGE BREAK
  //   row.addPageBreak();

  // VALIDATION
  worksheet.getCell("A1").dataValidation = {
    type: "list",
    allowBlank: true,
    formulae: ['"One,Two,Three,Four"'],
  };

  // WRITE TO FILE
  //   await workbook.xlsx.writeFile("test.xlsx");

  // WRITE TO BUFFER
  const wb = await workbook.xlsx.writeBuffer();
  const filename = "test.xlsx";
  reply.header("Content-disposition", "attachment; filename=" + filename);
  reply.header("Content-Type", "application/vnd.ms-excel");
  reply.send(wb);
});

// Run the server!
const start = async () => {
  try {
    await fastify.listen({ port: 3000 });
  } catch (err) {
    fastify.log.error(err);
    process.exit(1);
  }
};
start();

// getData();
