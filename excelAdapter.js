/* ExcelAdapter.js — Excel as the Database (Office.js) */
const ExcelAdapter = (() => {
  // Sheet & Table names
  const sheets = {
    vehicles:    { name: "Vehicles",    table: "tblVehicles",    headers: ["Plate","Make","Model","Year","Transmission","Rate","Status"] },
    customers:   { name: "Customers",   table: "tblCustomers",   headers: ["Name","Phone","Email","ID Type","ID No"] },
    rentals:     { name: "Rentals",     table: "tblRentals",     headers: ["Rental ID","Customer ID","Vehicle Plate","Start Date","Due Date","Actual Return","Daily Rate","Amount","Status"] },
    maintenance: { name: "Maintenance", table: "tblMaintenance", headers: ["Date","Vehicle Plate","Type","Odometer","Cost","Description"] },
  };

  /* ------------------------- Helpers ------------------------- */
  async function ensureWorksheet(context, name){
    const sheets = context.workbook.worksheets;
    try{
      sheets.getItem(name).load("name");
      await context.sync();
    }catch{
      const ws = sheets.add(name);
      ws.activate();
      await context.sync();
    }
  }
  async function ensureTable(context, sheetName, tableName, headers){
    const ws = context.workbook.worksheets.getItem(sheetName);
    const tables = context.workbook.tables;
    let table;
    try{
      table = tables.getItem(tableName);
      table.load("name");
      await context.sync();
    }catch{
      // Put headers at A1 row
      const hdrRange = ws.getRange("A1").getResizedRange(0, headers.length-1);
      hdrRange.values = [headers];
      hdrRange.format.font.bold = true;
      hdrRange.format.fill.color = "#CCEABB";
      // Create table
      table = tables.add(ws.getRangeByIndexes(0,0,1,headers.length), true /*hasHeaders*/);
      table.name = tableName;
      await context.sync();
    }
    return table;
  }

  function mapRowToObject(headers, row, rowIndexInData){
    const obj = {};
    headers.forEach((h,i)=> obj[toKey(h)] = row[i]);
    // Keep a rowId pointing to 0-based data row (excluding header)
    obj.__rowId = rowIndexInData;
    return obj;
  }
  function toKey(h){ return h.toLowerCase().replace(/\s+/g,''); }

  function val(v){ return (v===null||v===undefined)? "": v; }

  /* --------------------- Schema & Seed ---------------------- */
  async function ensureSchema(){
    await Excel.run(async (context)=>{
      // Create sheets if missing
      for(const s of Object.values(sheets)){
        await ensureWorksheet(context, s.name);
        await ensureTable(context, s.name, s.table, s.headers);
      }
    });
  }

  async function seedDemo(){
    await ensureSchema();
    await Excel.run(async (context)=>{
      // Vehicles
      const tv = context.workbook.tables.getItem(sheets.vehicles.table);
      const vRows = [
        ["NAB-1234","Toyota","Vios 1.3 E",2019,"AT",2000,"Available"],
        ["XYZ-5678","Honda","City 1.5",2021,"AT",2300,"Available"],
        ["AAA-1111","Mitsubishi","Mirage G4",2018,"MT",1700,"Maintenance"],
        ["BBB-2222","Toyota","Innova",2020,"AT",3500,"Reserved"],
      ];
      tv.rows.add(null, vRows);

      // Customers
      const tc = context.workbook.tables.getItem(sheets.customers.table);
      tc.rows.add(null, [
        ["Jane Doe","0917-000-1111","jane@example.com","DL","D-12345"],
        ["Juan Cruz","0918-222-3333","juan@example.com","DL","D-67890"],
        ["Maria S.","0917-888-9999","maria@example.com","Passport","P-556677"],
      ]);

      // Rentals
      const tr = context.workbook.tables.getItem(sheets.rentals.table);
      const today = new Date(); const iso = d=> d.toISOString().slice(0,10);
      const addDays = (d,n)=>{ const x=new Date(d); x.setDate(x.getDate()+n); return x; };
      tr.rows.add(null, [
        ["RENT-2025-001","Jane Doe","NAB-1234", iso(addDays(today,-2)), iso(addDays(today,4)), "", 2000, 2000*6, "Ongoing"],
        ["RENT-2025-002","Juan Cruz","XYZ-5678", iso(addDays(today,-35)), iso(addDays(today,-30)), iso(addDays(today,-30)), 2300, 2300*5, "Returned"],
      ]);

      // Maintenance
      const tm = context.workbook.tables.getItem(sheets.maintenance.table);
      tm.rows.add(null, [
        [ iso(new Date(Date.now()-15*86400000)), "AAA-1111", "Oil Change", 42000, 1800, "5W-30 full synthetic" ],
        [ iso(new Date(Date.now()-70*86400000)), "XYZ-5678", "Tire", 51000, 12000, "2 tires replaced" ],
      ]);

      await context.sync();
    });
  }

  /* ---------------------- Vehicles CRUD --------------------- */
  async function getVehicles(){
    await ensureSchema();
    return Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.vehicles.table);
      const rng = t.getDataBodyRange();
      rng.load(["values","rowCount","columnCount"]);
      const hdr = t.getHeaderRowRange(); hdr.load("values");
      await context.sync();

      const headers = hdr.values[0];
      if(rng.rowCount===0) return [];
      const out=[];
      for(let i=0;i<rng.rowCount;i++){
        const row = rng.values[i];
        out.push(mapRowToObject(headers, row, i));
      }
      return out;
    });
  }

  async function addVehicle(rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.vehicles.table);
      const row = [
        val(rec.plate), val(rec.make), val(rec.model),
        Number(rec.year||0), val(rec.transmission), Number(rec.rate||0),
        val(rec.status||"Available")
      ];
      t.rows.add(null, [row]);
      await context.sync();
    });
  }

  async function updateVehicle(rowId, rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.vehicles.table);
      // rowId is 0-based in data body
      const rng = t.getDataBodyRange().getRow(rowId);
      const vals = [
        val(rec.plate), val(rec.make), val(rec.model),
        Number(rec.year||0), val(rec.transmission), Number(rec.rate||0),
        val(rec.status||"Available")
      ];
      rng.values = [vals];
      await context.sync();
    });
  }

  async function deleteVehicle(rowId){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.vehicles.table);
      t.rows.getItemAt(rowId).delete();
      await context.sync();
    });
  }

  /* ---------------- TEMPLATES for other entities ------------- */
  // Copy these patterns and wire to your UI as needed.

  // Customers
  async function getCustomers(){
    await ensureSchema();
    return Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.customers.table);
      const body = t.getDataBodyRange(); const hdr = t.getHeaderRowRange();
      body.load(["values","rowCount"]); hdr.load("values");
      await context.sync();
      if(body.rowCount===0) return [];
      const headers = hdr.values[0]; const out=[];
      for(let i=0;i<body.rowCount;i++) out.push(mapRowToObject(headers, body.values[i], i));
      return out;
    });
  }
  async function addCustomer(rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.customers.table);
      t.rows.add(null, [[ val(rec.name), val(rec.phone), val(rec.email), val(rec.idType), val(rec.idNo) ]]);
      await context.sync();
    });
  }
  async function updateCustomer(rowId, rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.customers.table);
      const row=t.getDataBodyRange().getRow(rowId);
      row.values=[[ val(rec.name), val(rec.phone), val(rec.email), val(rec.idType), val(rec.idNo) ]];
      await context.sync();
    });
  }
  async function deleteCustomer(rowId){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.customers.table);
      t.rows.getItemAt(rowId).delete(); await context.sync();
    });
  }

  // Rentals
  async function getRentals(){
    await ensureSchema();
    return Excel.run(async (context)=>{
      const t = context.workbook.tables.getItem(sheets.rentals.table);
      const body=t.getDataBodyRange(); const hdr=t.getHeaderRowRange();
      body.load(["values","rowCount"]); hdr.load("values");
      await context.sync();
      if(body.rowCount===0) return [];
      const headers = hdr.values[0]; const out=[];
      for(let i=0;i<body.rowCount;i++) out.push(mapRowToObject(headers, body.values[i], i));
      return out;
    });
  }
  async function addRental(rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.rentals.table);
      t.rows.add(null, [[
        val(rec.rentalid||rec["rentalid"]||rec["Rental ID"]||rec.rentalId),
        val(rec.customerid||rec["Customer ID"]||rec.customer||rec.customerName||""),
        val(rec.vehicleplate||rec["Vehicle Plate"]||rec.vehicle||rec.vehiclePlate||""),
        val(rec.startdate||rec["Start Date"]||rec.startDate||""),
        val(rec.duedate||rec["Due Date"]||rec.dueDate||""),
        val(rec.actualreturn||rec["Actual Return"]||rec.actualReturn||""),
        Number(rec.dailyrate||rec["Daily Rate"]||rec.dailyRate||0),
        Number(rec.amount||rec["Amount"]||0),
        val(rec.status||"Ongoing")
      ]]);
      await context.sync();
    });
  }
  async function updateRental(rowId, rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.rentals.table);
      const row=t.getDataBodyRange().getRow(rowId);
      row.values=[[
        val(rec.rentalId),
        val(rec.customerId||rec.customerName||""),
        val(rec.vehiclePlate||""),
        val(rec.startDate||""),
        val(rec.dueDate||""),
        val(rec.actualReturn||""),
        Number(rec.dailyRate||0),
        Number(rec.amount||0),
        val(rec.status||"Ongoing")
      ]];
      await context.sync();
    });
  }
  async function deleteRental(rowId){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.rentals.table);
      t.rows.getItemAt(rowId).delete(); await context.sync();
    });
  }

  // Maintenance
  async function getMaintenance(){
    await ensureSchema();
    return Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.maintenance.table);
      const body=t.getDataBodyRange(); const hdr=t.getHeaderRowRange();
      body.load(["values","rowCount"]); hdr.load("values");
      await context.sync();
      if(body.rowCount===0) return [];
      const headers = hdr.values[0]; const out=[];
      for(let i=0;i<body.rowCount;i++) out.push(mapRowToObject(headers, body.values[i], i));
      return out;
    });
  }
  async function addMaintenance(rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.maintenance.table);
      t.rows.add(null, [[
        val(rec.date||""), val(rec.vehiclePlate||""), val(rec.type||""),
        Number(rec.odometer||0), Number(rec.cost||0), val(rec.description||rec.desc||"")
      ]]);
      await context.sync();
    });
  }
  async function updateMaintenance(rowId, rec){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.maintenance.table);
      const row=t.getDataBodyRange().getRow(rowId);
      row.values=[[
        val(rec.date||""), val(rec.vehiclePlate||""), val(rec.type||""),
        Number(rec.odometer||0), Number(rec.cost||0), val(rec.description||rec.desc||"")
      ]];
      await context.sync();
    });
  }
  async function deleteMaintenance(rowId){
    await ensureSchema();
    await Excel.run(async (context)=>{
      const t=context.workbook.tables.getItem(sheets.maintenance.table);
      t.rows.getItemAt(rowId).delete(); await context.sync();
    });
  }

  /* -------------------- Public API -------------------------- */
  return {
    ensureSchema, seedDemo,
    // Vehicles
    getVehicles, addVehicle, updateVehicle, deleteVehicle,
    // Customers
    getCustomers, addCustomer, updateCustomer, deleteCustomer,
    // Rentals
    getRentals, addRental, updateRental, deleteRental,
    // Maintenance
    getMaintenance, addMaintenance, updateMaintenance, deleteMaintenance,
  };
})();
