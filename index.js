// read excel .xls

const XLSX = require("xlsx");

const serialize = (data) => {
  return data?.map((d) => {
    const trim = (str) => str?.replace(/^\s+|\s+$/g, "");

    return {
      id: d["ID"],
      nama: trim(d["NAMA_UNOR"]),
      eselon_id: d["ESELON_ID"],
      cepat_kode: d["CEPAT_KODE"],
      diatasan_id: d["DIATASAN_ID"],
      instansi_id: d["INSTANSI_ID"],
      pemimpin_pns_id: d["PEMIMPIN_PNS_ID"],
      unor_induk_id: d["UNOR_INDUK"],
    };
  });
};

(async () => {
  const workbook = XLSX.readFile("./hierarchicalUnor.xls");
  //   read first sheet
  const sheet_name_list = workbook.SheetNames;

  //   read 5th rows as header and 6th rows as data
  const xlData2 = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

  const data = serialize(xlData2);

  //   make insert query
  const insertQuery = data.map((d) => {
    return `INSERT INTO unor (id, nama, eselon_id, cepat_kode, diatasan_id, instansi_id, pemimpin_pns_id, unor_induk_id) VALUES ('${d.id}', '${d.nama}', '${d.eselon_id}', '${d.cepat_kode}', '${d.diatasan_id}', '${d.instansi_id}', '${d.pemimpin_pns_id}', '${d.unor_induk_id}');`;
  });

  //   save as .sql file
  const fs = require("fs");
  fs.writeFile("unor.sql", insertQuery.join("\n"), function (err) {
    if (err) return console.log(err);
    console.log("unor.sql saved!");
  });
})();
