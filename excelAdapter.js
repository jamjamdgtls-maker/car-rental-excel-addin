/* ExcelAdapter.js — Excel as the Database (Office.js) */
const ExcelAdapter = (() => {
  const sheets = {
    vehicles:    { name: "Vehicles",    table: "tblVehicles",    headers: ["Plate","Make","Model","Year","Transmission","Rate","Status"] },
    customers:   { name: "Customers",   table: "tblCustomers",   headers: ["Name","Phone","Email","ID Type","ID No"] },
    rentals:     { name: "Rentals",     table: "tblRentals",     headers: ["Rental ID","Customer ID","Vehicle Plate","Start Date","Due Date","Actual Return","Daily Rate","Amount","Status"] },
    maintenance: { name: "Maintenance", table: "tblMaintenance", headers: ["Date","Vehicle Plate","Type","Odometer","Cost","Description"] },
  };

  async function ensureWorksheet(context, name){
    const wsColl = context.workbook.worksheets;
    try{ wsColl.getItem(name).load("name"); await context.sync(); }
    catch{ wsColl.add(name); await context.sync(); }
  }
  async function ensureTable(context, sheetName, tableName, headers){
    const ws = context.workbook.worksheets.getItem(sheetName);
    const tables = context.workbook.tables;
    let table;
    try{ table=tables.getItem(tableName); table.load("name"); await context.sync(); }
    catch{
      const hdrRange = ws.getRange("A1").getResizedRange(0, headers.length-1);
      hdrRange.values = [headers];
      hdrRange.format.font.bold = true;
      hdrRange.format.fill.color = "#CCEABB";
      table = tables.add(ws.getRangeByIndexes(0,0,1,headers.length), true);
      table.name = tableName;
      await context.sync();
    }
    return table;
  }
  function mapRowToObject(headers, row, i){ const o={}; headers.forEach((h,idx)=>o[toKey(h)]=row[idx]); o.__rowId=i; return o; }
  const toKey = h => h.toLowerCase().replace(/\s+/g,'');
  const val = v => (v==null? "" : v);

  async function ensureSchema(){
    await Excel.run(async (context)=>{
      for(const s of Object.values(sheets)){
        await ensureWorksheet(context, s.name);
        await ensureTable(context, s.name, s.table, s.headers);
      }
    });
  }

  async function seedDemo(){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const today = new Date(); const iso = d=> d.toISOString().slice(0,10);
      const addDays = (d,n)=>{ const x=new Date(d); x.setDate(x.getDate()+n); return x; };

      context.workbook.tables.getItem(sheets.vehicles.table).rows.add(null, [
        ["NAB-1234","Toyota","Vios 1.3 E",2019,"AT",2000,"Available"],
        ["XYZ-5678","Honda","City 1.5",2021,"AT",2300,"Available"],
        ["AAA-1111","Mitsubishi","Mirage G4",2018,"MT",1700,"Maintenance"],
        ["BBB-2222","Toyota","Innova",2020,"AT",3500,"Reserved"],
      ]);

      context.workbook.tables.getItem(sheets.customers.table).rows.add(null, [
        ["Jane Doe","0917-000-1111","jane@example.com","DL","D-12345"],
        ["Juan Cruz","0918-222-3333","juan@example.com","DL","D-67890"],
        ["Maria S.","0917-888-9999","maria@example.com","Passport","P-556677"],
      ]);

      context.workbook.tables.getItem(sheets.rentals.table).rows.add(null, [
        ["RENT-2025-001","Jane Doe","NAB-1234", iso(addDays(today,-2)), iso(addDays(today,4)), "", 2000, 2000*6, "Ongoing"],
        ["RENT-2025-002","Juan Cruz","XYZ-5678", iso(addDays(today,-35)), iso(addDays(today,-30)), iso(addDays(today,-30)), 2300, 2300*5, "Returned"],
      ]);

      context.workbook.tables.getItem(sheets.maintenance.table).rows.add(null, [
        [ iso(new Date(Date.now()-15*86400000)), "AAA-1111", "Oil Change", 42000, 1800, "5W-30 full synthetic" ],
        [ iso(new Date(Date.now()-70*86400000)), "XYZ-5678", "Tire", 51000, 12000, "2 tires replaced" ],
      ]);

      await context.sync();
    });
  }

  // VEHICLES CRUD
  async function getVehicles(){
    await ensureSchema();
    return Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.vehicles.table);
      const body = t.getDataBodyRange(); const hdr=t.getHeaderRowRange();
      body.load(["values","rowCount"]); hdr.load("values"); await context.sync();
      if(body.rowCount===0) return [];
      const headers = hdr.values[0]; const out=[];
      for(let i=0;i<body.rowCount;i++) out.push(mapRowToObject(headers, body.values[i], i));
      return out;
    });
  }
  async function addVehicle(rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.vehicles.table);
      t.rows.add(null, [[
        val(rec.plate), val(rec.make), val(rec.model),
        Number(rec.year||0), val(rec.transmission), Number(rec.rate||0), val(rec.status||"Available")
      ]]); await context.sync();
    });
  }
  async function updateVehicle(rowId, rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.vehicles.table);
      t.getDataBodyRange().getRow(rowId).values=[[
        val(rec.plate), val(rec.make), val(rec.model),
        Number(rec.year||0), val(rec.transmission), Number(rec.rate||0), val(rec.status||"Available")
      ]]; await context.sync();
    });
  }
  async function deleteVehicle(rowId){
    await ensureSchema();
    await Excel.run(async (context)=>{
      context.workbook.tables.getItem(sheets.vehicles.table).rows.getItemAt(rowId).delete();
      await context.sync();
    });
  }

  return {
    ensureSchema, seedDemo,
    getVehicles, addVehicle, updateVehicle, deleteVehicle,
  };
})();
