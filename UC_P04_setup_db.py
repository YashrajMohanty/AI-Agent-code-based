"""
UC-P04 | Procurement AI Agent — Database Setup
Run FIRST to create procurement.db SQLite database.
"""
import sqlite3, os, random
from datetime import date, timedelta

DB=os.path.join(os.path.dirname(__file__),"UC_P04_procurement.db")
conn=sqlite3.connect(DB)
c=conn.cursor()

for table_name in ["vendors","historical_pos","benchmark_rates"]:
    c.execute(f"DROP TABLE IF EXISTS {table_name}")

c.execute("""CREATE TABLE vendors(vendor_id TEXT PRIMARY KEY,name TEXT,category TEXT,country TEXT,rating REAL,lead_days INTEGER)""")
c.execute("""CREATE TABLE historical_pos(po_id TEXT,vendor_id TEXT,item TEXT,unit TEXT,qty REAL,price_inr REAL,order_date TEXT,delivery_days INTEGER,status TEXT)""")
c.execute("""CREATE TABLE benchmark_rates(item TEXT PRIMARY KEY,unit TEXT,market_rate_inr REAL,updated_date TEXT)""")

vendors=[
    ("V001","Hindalco Chemicals","Caustic Soda","India",4.2,21),
    ("V002","Indo Gulf Corp","Caustic Soda","India",3.8,18),
    ("V003","Gulf Chem Trading","Caustic Soda","UAE",3.5,45),
    ("V004","Chengdu Carbon","Carbon Anodes","China",3.9,60),
    ("V005","Gulf Anode Co","Carbon Anodes","UAE",4.1,42),
    ("V006","CPCL Suppliers","HFO","India",4.5,14),
    ("V007","SRF Industries","Aluminium Fluoride","India",4.3,25),
    ("V008","Bhushan Metals","Structural Steel","India",3.7,20),
    ("V009","L&T Construction","Civil Works","India",4.6,30),
    ("V010","Adani Logistics","Logistics","India",4.0,7),
]
c.executemany("INSERT INTO vendors VALUES(?,?,?,?,?,?)",vendors)

items_map={"Caustic Soda":("MT",38000),
           "Carbon Anodes":("MT",320000),
           "HFO":("KL",52000),
           "Aluminium Fluoride":("MT",95000),
           "Structural Steel":("MT",68000),
           "Civil Works":("LS",500000),
           "Logistics":("Trip",15000)}

random.seed(42); po_id=1
for _,(vid,name,cat,*_) in enumerate(vendors):
    for _ in range(random.randint(5,12)):
        d=date(2023,1,1)+timedelta(days=random.randint(0,700))
        unit,base=items_map.get(cat,("Unit",50000))
        price=round(base*(1+random.gauss(0,0.08)),0)
        qty=round(random.uniform(50,600),1)
        c.execute("INSERT INTO historical_pos VALUES(?,?,?,?,?,?,?,?,?)",
                  (f"PO-{po_id:05d}",vid,cat,unit,qty,price,str(d),
                   random.randint(14,75),random.choice(["Delivered","Delivered","Delivered","Delayed"])))
        po_id+=1

benchmarks=[
    ("Caustic Soda","MT",36500,str(date.today())),
    ("Carbon Anodes","MT",315000,str(date.today())),
    ("HFO","KL",49000,str(date.today())),
    ("Aluminium Fluoride","MT",92000,str(date.today())),
    ("Structural Steel","MT",65000,str(date.today())),
    ("Logistics","Trip",14500,str(date.today())),
]
c.executemany("INSERT INTO benchmark_rates VALUES(?,?,?,?)",benchmarks)

conn.commit()
conn.close()

print(f"UC-P04 database created: {DB} with {po_id-1} historical POs.")