const fastify = require("fastify")({ logger: true });

var XLSX = require("xlsx");
const fs = require("fs");
const stream = require("stream");
XLSX.stream.set_readable(stream.Readable);
XLSX.set_fs(fs);

// Declare a route
fastify.get("/", async (request, reply) => {
  // var fileContents = Buffer.from(fileData, "base64");

  const mockData = [
    {
      name: "yum",
      id: 1,
    },
    {
      name: "haha",
      id: 2,
    },
  ];

  // available utilities method
  // https://docs.sheetjs.com/docs/api/#utilities

  // possible file format
  // https://docs.sheetjs.com/docs/miscellany/formats

  //create workbook
  const workbook = XLSX.utils.book_new();

  //create worksheet with data
  const worksheet = XLSX.utils.json_to_sheet(mockData);

  // test modify header, now name, id, will change to Name, Birthday
  XLSX.utils.sheet_add_aoa(worksheet, [["Name", "Birthday"]], { origin: "A1" });

  // change column width
  const max_width = mockData.reduce((w, r) => Math.max(w, r.name.length), 10);
  worksheet["!cols"] = [{ wch: max_width }];

  // add worksheet to workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, "sample");

  // add another worksheet with aoa data format for testing
  var ws_data = [
    ["S", "h", "e", "e", "t", "J", "S"],
    [1, 2, 3, 4, 5],
  ];

  var ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(workbook, ws, "aoa");

  // write workbook to file
  // XLSX.writeFile(workbook, "sample.xlsx", { compression: true });

  // write to buffer, possible options 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';
  const wb = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  // async
  // example how read sheets data
  console.log(workbook.Sheets.sample);

  const filename = "test.xlsx";
  reply.header("Content-disposition", "attachment; filename=" + filename);
  reply.header("Content-Type", "application/vnd.ms-excel");
  // "application/vnd.ms-excel" "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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
