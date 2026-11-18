import os
import sqlite3
import datetime
from typing import Optional, List, Dict, Any
from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "legal_register.db")
EXCEL_FILE = os.path.join(BASE_DIR, "Legal_Register.xlsx")

app = FastAPI(title="Legal Register & Compliance System", version="1.0.0")

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
    CREATE TABLE IF NOT EXISTS laws (
        law_id TEXT PRIMARY KEY,
        title TEXT NOT NULL,
        law_type TEXT,
        ministry TEXT,
        effective_date TEXT,
        status TEXT DEFAULT 'ใช้อยู่',
        year_added INTEGER,
        source_link TEXT
    )''')

    cur.execute('''
    CREATE TABLE IF NOT EXISTS requirements (
        req_id TEXT PRIMARY KEY,
        law_id TEXT NOT NULL,
        section TEXT,
        summary TEXT NOT NULL,
        department TEXT,
        FOREIGN KEY(law_id) REFERENCES laws(law_id)
    )''')

    cur.execute('''
    CREATE TABLE IF NOT EXISTS assessments (
        assess_id TEXT PRIMARY KEY,
        req_id TEXT NOT NULL,
        assessment_date TEXT NOT NULL,
        assessor TEXT NOT NULL,
        result TEXT NOT NULL,
        evidence TEXT,
        action_plan TEXT,
        due_date TEXT,
        FOREIGN KEY(req_id) REFERENCES requirements(req_id)
    )''')

    cur.execute('''
    CREATE TABLE IF NOT EXISTS change_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        log_date TEXT,
        log_year INTEGER,
        action_type TEXT,
        item_id TEXT,
        description TEXT,
        author TEXT
    )''')
    conn.commit()

    # Check if empty, seed from Excel
    cur.execute("SELECT COUNT(*) FROM laws")
    if cur.fetchone()[0] == 0 and os.path.exists(EXCEL_FILE):
        seed_from_excel(conn)
    conn.close()

def seed_from_excel(conn):
    try:
        cur = conn.cursor()
        wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
        
        if 'Laws' in wb.sheetnames:
            ws = wb['Laws']
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    cur.execute(
                        "INSERT OR IGNORE INTO laws VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        (row[0], row[1], row[2], row[3], str(row[4]) if row[4] else None, row[5] or 'ใช้อยู่', row[6], row[7])
                    )
                    
        if 'Requirements' in wb.sheetnames:
            ws = wb['Requirements']
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    cur.execute(
                        "INSERT OR IGNORE INTO requirements VALUES (?, ?, ?, ?, ?)",
                        (row[0], row[1], row[3], row[4], row[5])
                    )
                    
        if 'Assessments' in wb.sheetnames:
            ws = wb['Assessments']
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    ad = row[3].strftime('%Y-%m-%d') if isinstance(row[3], datetime.datetime) else (str(row[3]) if row[3] else None)
                    dd = row[8].strftime('%Y-%m-%d') if isinstance(row[8], datetime.datetime) else (str(row[8]) if row[8] else None)
                    cur.execute(
                        "INSERT OR IGNORE INTO assessments VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        (row[0], row[1], ad, row[4], row[5], row[6], row[7], dd)
                    )
                    
        if 'Change_Log' in wb.sheetnames:
            ws = wb['Change_Log']
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] or row[2]:
                    ld = row[0].strftime('%Y-%m-%d') if isinstance(row[0], datetime.datetime) else (str(row[0]) if row[0] else None)
                    cur.execute(
                        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
                        (ld, row[1], row[2], row[3], row[4], row[5])
                    )
        conn.commit()
        print("Data successfully seeded from Excel!")
    except Exception as e:
        print(f"Error seeding data: {e}")

init_db()

# Models
class LawCreate(BaseModel):
    title: str
    law_type: str
    ministry: str
    effective_date: Optional[str] = None
    status: Optional[str] = "ใช้อยู่"
    year_added: Optional[int] = None
    source_link: Optional[str] = None
    author: Optional[str] = "Admin"

class LawUpdate(BaseModel):
    title: str
    law_type: str
    ministry: str
    effective_date: Optional[str] = None
    status: str
    year_added: Optional[int] = None
    source_link: Optional[str] = None
    author: Optional[str] = "Admin"

class ReqCreate(BaseModel):
    law_id: str
    section: Optional[str] = None
    summary: str
    department: str
    author: Optional[str] = "Admin"

class ReqUpdate(BaseModel):
    law_id: str
    section: Optional[str] = None
    summary: str
    department: str
    author: Optional[str] = "Admin"

class AssessmentCreate(BaseModel):
    req_id: str
    assessment_date: str
    assessor: str
    result: str
    evidence: Optional[str] = "-"
    action_plan: Optional[str] = "-"
    due_date: Optional[str] = None

class ChangeLogCreate(BaseModel):
    action_type: str
    item_id: str
    description: str
    author: str

# Helper functions for next IDs
def get_next_id(table: str, prefix: str) -> str:
    conn = get_db()
    cur = conn.cursor()
    col = f"{prefix.lower()}_id" if prefix.lower() != "a" else "assess_id"
    if table == "laws":
        col = "law_id"
    elif table == "requirements":
        col = "req_id"
    elif table == "assessments":
        col = "assess_id"

    cur.execute(f"SELECT {col} FROM {table}")
    rows = cur.fetchall()
    conn.close()
    
    max_num = 0
    for r in rows:
        val = r[0]
        if val and val.startswith(prefix):
            try:
                num = int(val[len(prefix):])
                if num > max_num:
                    max_num = num
            except ValueError:
                pass
    return f"{prefix}{str(max_num + 1).zfill(3)}"

# API Endpoints
@app.get("/api/dashboard")
def get_dashboard_data():
    conn = get_db()
    cur = conn.cursor()
    today_str = datetime.date.today().strftime('%Y-%m-%d')

    # Total active laws
    cur.execute("SELECT COUNT(*) FROM laws WHERE status = 'ใช้อยู่'")
    active_laws = cur.fetchone()[0]

    # Total laws all
    cur.execute("SELECT COUNT(*) FROM laws")
    total_laws = cur.fetchone()[0]

    # Total requirements
    cur.execute("SELECT COUNT(*) FROM requirements")
    total_requirements = cur.fetchone()[0]

    # Latest assessment for each requirement
    cur.execute('''
    SELECT 
        r.req_id,
        r.law_id,
        l.title as law_title,
        r.section,
        r.summary,
        r.department,
        a.assess_id,
        a.assessment_date,
        a.assessor,
        a.result,
        a.action_plan,
        a.due_date,
        a.evidence
    FROM requirements r
    LEFT JOIN laws l ON r.law_id = l.law_id
    LEFT JOIN (
        SELECT a1.*
        FROM assessments a1
        JOIN (
            SELECT req_id, MAX(assessment_date) as max_date, MAX(assess_id) as max_id
            FROM assessments
            GROUP BY req_id
        ) a2 ON a1.req_id = a2.req_id AND a1.assessment_date = a2.max_date AND a1.assess_id = a2.max_id
    ) a ON r.req_id = a.req_id
    ''')
    req_status_rows = cur.fetchall()

    status_counts = {
        "สอดคล้อง": 0,
        "ไม่สอดคล้อง": 0,
        "ไม่เกี่ยวข้อง": 0,
        "ยังไม่ประเมิน": 0
    }
    
    dept_stats = {}
    overdue_items = []
    latest_assess_date = None

    for row in req_status_rows:
        res = row["result"]
        dept = row["department"] or "ไม่ระบุ"
        if dept not in dept_stats:
            dept_stats[dept] = {"total": 0, "compliant": 0, "non_compliant": 0, "na": 0, "unassessed": 0}
        dept_stats[dept]["total"] += 1

        if not res:
            status_counts["ยังไม่ประเมิน"] += 1
            dept_stats[dept]["unassessed"] += 1
        else:
            if res in status_counts:
                status_counts[res] += 1
            else:
                status_counts[res] = 1

            if res == "สอดคล้อง":
                dept_stats[dept]["compliant"] += 1
            elif res == "ไม่สอดคล้อง":
                dept_stats[dept]["non_compliant"] += 1
            elif res == "ไม่เกี่ยวข้อง":
                dept_stats[dept]["na"] += 1

        # Check latest date
        ad = row["assessment_date"]
        if ad:
            if latest_assess_date is None or ad > latest_assess_date:
                latest_assess_date = ad

        # Check overdue
        if res == "ไม่สอดคล้อง":
            due_date = row["due_date"]
            if due_date and due_date < today_str:
                days_overdue = (datetime.datetime.strptime(today_str, '%Y-%m-%d') - datetime.datetime.strptime(due_date, '%Y-%m-%d')).days
                overdue_items.append({
                    "req_id": row["req_id"],
                    "law_id": row["law_id"],
                    "law_title": row["law_title"],
                    "section": row["section"],
                    "summary": row["summary"],
                    "department": row["department"],
                    "assess_id": row["assess_id"],
                    "assessment_date": row["assessment_date"],
                    "assessor": row["assessor"],
                    "action_plan": row["action_plan"],
                    "due_date": due_date,
                    "days_overdue": days_overdue
                })

    # Compliance Rate
    compliant = status_counts["สอดคล้อง"]
    non_compliant = status_counts["ไม่สอดคล้อง"]
    eval_base = compliant + non_compliant
    compliance_rate = round((compliant / eval_base * 100), 1) if eval_base > 0 else 0.0

    # Yearly trends
    cur.execute('''
    SELECT 
        strftime('%Y', assessment_date) as year,
        COUNT(*) as total_assessments,
        SUM(CASE WHEN result = 'สอดคล้อง' THEN 1 ELSE 0 END) as compliant,
        SUM(CASE WHEN result = 'ไม่สอดคล้อง' THEN 1 ELSE 0 END) as non_compliant,
        SUM(CASE WHEN result = 'ไม่เกี่ยวข้อง' THEN 1 ELSE 0 END) as not_applicable
    FROM assessments
    WHERE assessment_date IS NOT NULL
    GROUP BY strftime('%Y', assessment_date)
    ORDER BY year ASC
    ''')
    yearly_rows = cur.fetchall()

    # Yearly new laws count
    cur.execute('''
    SELECT log_year, COUNT(*) as new_laws
    FROM change_log
    WHERE action_type = 'เพิ่ม' AND item_id LIKE 'L%'
    GROUP BY log_year
    ''')
    yearly_laws_map = {row["log_year"]: row["new_laws"] for row in cur.fetchall()}

    yearly_data = []
    for y in yearly_rows:
        yr = int(y["year"]) if y["year"] and y["year"].isdigit() else y["year"]
        c = y["compliant"] or 0
        nc = y["non_compliant"] or 0
        tot = c + nc
        rate = round((c / tot * 100), 1) if tot > 0 else 0.0
        yearly_data.append({
            "year": yr,
            "total": y["total_assessments"],
            "compliant": c,
            "non_compliant": nc,
            "not_applicable": y["not_applicable"] or 0,
            "compliance_rate": rate,
            "new_laws": yearly_laws_map.get(yr, 0)
        })

    conn.close()

    return {
        "metrics": {
            "active_laws": active_laws,
            "total_laws": total_laws,
            "total_requirements": total_requirements,
            "compliance_rate": compliance_rate,
            "compliant_count": compliant,
            "non_compliant_count": non_compliant,
            "na_count": status_counts["ไม่เกี่ยวข้อง"],
            "unassessed_count": status_counts["ยังไม่ประเมิน"],
            "overdue_count": len(overdue_items),
            "latest_assessment_date": latest_assess_date or "-"
        },
        "status_distribution": status_counts,
        "department_stats": dept_stats,
        "yearly_trends": yearly_data,
        "overdue_items": overdue_items
    }

@app.get("/api/laws")
def list_laws(status: Optional[str] = None, law_type: Optional[str] = None, search: Optional[str] = None):
    conn = get_db()
    cur = conn.cursor()
    query = '''
    SELECT 
        l.*,
        COUNT(r.req_id) as req_count
    FROM laws l
    LEFT JOIN requirements r ON l.law_id = r.law_id
    WHERE 1=1
    '''
    params = []
    if status:
        query += " AND l.status = ?"
        params.append(status)
    if law_type:
        query += " AND l.law_type = ?"
        params.append(law_type)
    if search:
        query += " AND (l.title LIKE ? OR l.law_id LIKE ? OR l.ministry LIKE ?)"
        wildcard = f"%{search}%"
        params.extend([wildcard, wildcard, wildcard])

    query += " GROUP BY l.law_id ORDER BY l.law_id ASC"
    cur.execute(query, params)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@app.get("/api/laws/next-id")
def get_next_law_id():
    return {"next_id": get_next_id("laws", "L")}

@app.post("/api/laws")
def create_law(data: LawCreate):
    conn = get_db()
    cur = conn.cursor()
    law_id = get_next_id("laws", "L")
    year_added = data.year_added or datetime.date.today().year

    cur.execute(
        "INSERT INTO laws (law_id, title, law_type, ministry, effective_date, status, year_added, source_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (law_id, data.title, data.law_type, data.ministry, data.effective_date, data.status, year_added, data.source_link)
    )

    # Log to Change_Log
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today_str, year_added, "เพิ่ม", law_id, f"เพิ่มกฎหมายใหม่: {data.title}", data.author or "Admin")
    )
    conn.commit()
    conn.close()
    return {"message": "Law created successfully", "law_id": law_id}

@app.put("/api/laws/{law_id}")
def update_law(law_id: str, data: LawUpdate):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM laws WHERE law_id = ?", (law_id,))
    existing = cur.fetchone()
    if not existing:
        conn.close()
        raise HTTPException(status_code=404, detail="Law not found")

    cur.execute(
        "UPDATE laws SET title = ?, law_type = ?, ministry = ?, effective_date = ?, status = ?, year_added = ?, source_link = ? WHERE law_id = ?",
        (data.title, data.law_type, data.ministry, data.effective_date, data.status, data.year_added, data.source_link, law_id)
    )

    # Log change
    today = datetime.date.today()
    action = "ยกเลิก" if data.status == "ยกเลิก" else "แก้ไข"
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today.strftime('%Y-%m-%d'), today.year, action, law_id, f"แก้ไขกฎหมาย: {data.title} (สถานะ: {data.status})", data.author or "Admin")
    )
    conn.commit()
    conn.close()
    return {"message": "Law updated successfully"}

@app.delete("/api/laws/{law_id}")
def delete_law(law_id: str, author: Optional[str] = "Admin"):
    conn = get_db()
    cur = conn.cursor()
    # Check if has requirements
    cur.execute("SELECT COUNT(*) FROM requirements WHERE law_id = ?", (law_id,))
    req_count = cur.fetchone()[0]
    if req_count > 0:
        conn.close()
        raise HTTPException(status_code=400, detail=f"Cannot delete law that has {req_count} requirements attached. Please delete requirements first or set law status to 'ยกเลิก'.")

    cur.execute("DELETE FROM laws WHERE law_id = ?", (law_id,))
    today = datetime.date.today()
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today.strftime('%Y-%m-%d'), today.year, "ยกเลิก", law_id, f"ลบกฎหมายรหัส {law_id}", author or "Admin")
    )
    conn.commit()
    conn.close()
    return {"message": "Law deleted successfully"}

# Requirements
@app.get("/api/requirements")
def list_requirements(law_id: Optional[str] = None, department: Optional[str] = None, result: Optional[str] = None, search: Optional[str] = None):
    conn = get_db()
    cur = conn.cursor()
    query = '''
    SELECT 
        r.*,
        l.title as law_title,
        l.status as law_status,
        a.assess_id as latest_assess_id,
        a.assessment_date as latest_assessment_date,
        a.assessor as latest_assessor,
        COALESCE(a.result, 'ยังไม่ประเมิน') as latest_result,
        a.evidence as latest_evidence,
        a.action_plan as latest_action_plan,
        a.due_date as latest_due_date
    FROM requirements r
    JOIN laws l ON r.law_id = l.law_id
    LEFT JOIN (
        SELECT a1.*
        FROM assessments a1
        JOIN (
            SELECT req_id, MAX(assessment_date) as max_date, MAX(assess_id) as max_id
            FROM assessments
            GROUP BY req_id
        ) a2 ON a1.req_id = a2.req_id AND a1.assessment_date = a2.max_date AND a1.assess_id = a2.max_id
    ) a ON r.req_id = a.req_id
    WHERE 1=1
    '''
    params = []
    if law_id:
        query += " AND r.law_id = ?"
        params.append(law_id)
    if department:
        query += " AND r.department LIKE ?"
        params.append(f"%{department}%")
    if result:
        if result == "ยังไม่ประเมิน":
            query += " AND a.result IS NULL"
        else:
            query += " AND a.result = ?"
            params.append(result)
    if search:
        query += " AND (r.summary LIKE ? OR r.section LIKE ? OR r.req_id LIKE ? OR l.title LIKE ?)"
        wildcard = f"%{search}%"
        params.extend([wildcard, wildcard, wildcard, wildcard])

    query += " ORDER BY r.req_id ASC"
    cur.execute(query, params)
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@app.get("/api/requirements/next-id")
def get_next_req_id():
    return {"next_id": get_next_id("requirements", "R")}

@app.post("/api/requirements")
def create_requirement(data: ReqCreate):
    conn = get_db()
    cur = conn.cursor()
    req_id = get_next_id("requirements", "R")

    cur.execute(
        "INSERT INTO requirements (req_id, law_id, section, summary, department) VALUES (?, ?, ?, ?, ?)",
        (req_id, data.law_id, data.section or "-", data.summary, data.department)
    )

    today = datetime.date.today()
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today.strftime('%Y-%m-%d'), today.year, "เพิ่ม", req_id, f"เพิ่มข้อกำหนด: {data.summary} (ผูกกับ {data.law_id})", data.author or "Admin")
    )
    conn.commit()
    conn.close()
    return {"message": "Requirement created successfully", "req_id": req_id}

@app.put("/api/requirements/{req_id}")
def update_requirement(req_id: str, data: ReqUpdate):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "UPDATE requirements SET law_id = ?, section = ?, summary = ?, department = ? WHERE req_id = ?",
        (data.law_id, data.section, data.summary, data.department, req_id)
    )
    today = datetime.date.today()
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today.strftime('%Y-%m-%d'), today.year, "แก้ไข", req_id, f"ปรับปรุงข้อกำหนด: {data.summary}", data.author or "Admin")
    )
    conn.commit()
    conn.close()
    return {"message": "Requirement updated successfully"}

@app.delete("/api/requirements/{req_id}")
def delete_requirement(req_id: str, author: Optional[str] = "Admin"):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM assessments WHERE req_id = ?", (req_id,))
    cur.execute("DELETE FROM requirements WHERE req_id = ?", (req_id,))
    today = datetime.date.today()
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today.strftime('%Y-%m-%d'), today.year, "ยกเลิก", req_id, f"ลบข้อกำหนดรหัส {req_id}", author or "Admin")
    )
    conn.commit()
    conn.close()
    return {"message": "Requirement and associated assessments deleted"}

# Assessments
@app.get("/api/assessments")
def list_assessments(req_id: Optional[str] = None):
    conn = get_db()
    cur = conn.cursor()
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    query = '''
    SELECT 
        a.*,
        r.section,
        r.summary as req_summary,
        r.department,
        l.law_id,
        l.title as law_title
    FROM assessments a
    JOIN requirements r ON a.req_id = r.req_id
    JOIN laws l ON r.law_id = l.law_id
    WHERE 1=1
    '''
    params = []
    if req_id:
        query += " AND a.req_id = ?"
        params.append(req_id)
    query += " ORDER BY a.assessment_date DESC, a.assess_id DESC"
    cur.execute(query, params)
    
    rows = []
    for r in cur.fetchall():
        d = dict(r)
        is_overdue = False
        if d["result"] == "ไม่สอดคล้อง" and d["due_date"]:
            if d["due_date"] < today_str:
                is_overdue = True
        d["is_overdue"] = is_overdue
        rows.append(d)
    conn.close()
    return rows

@app.get("/api/assessments/next-id")
def get_next_assess_id():
    return {"next_id": get_next_id("assessments", "A")}

@app.post("/api/assessments")
def create_assessment(data: AssessmentCreate):
    conn = get_db()
    cur = conn.cursor()
    assess_id = get_next_id("assessments", "A")

    cur.execute(
        "INSERT INTO assessments (assess_id, req_id, assessment_date, assessor, result, evidence, action_plan, due_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (assess_id, data.req_id, data.assessment_date, data.assessor, data.result, data.evidence or "-", data.action_plan or "-", data.due_date)
    )

    # If action plan and overdue or non-compliant, log
    today = datetime.date.today()
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today.strftime('%Y-%m-%d'), today.year, "ประเมิน", data.req_id, f"บันทึกผลการประเมิน: {data.result} (รอบ {data.assessment_date})", data.assessor)
    )

    conn.commit()
    conn.close()
    return {"message": "Assessment recorded successfully", "assess_id": assess_id}

@app.put("/api/assessments/{assess_id}")
def update_assessment(assess_id: str, data: AssessmentCreate):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "UPDATE assessments SET req_id = ?, assessment_date = ?, assessor = ?, result = ?, evidence = ?, action_plan = ?, due_date = ? WHERE assess_id = ?",
        (data.req_id, data.assessment_date, data.assessor, data.result, data.evidence or "-", data.action_plan or "-", data.due_date, assess_id)
    )
    conn.commit()
    conn.close()
    return {"message": "Assessment updated successfully"}

@app.delete("/api/assessments/{assess_id}")
def delete_assessment(assess_id: str):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM assessments WHERE assess_id = ?", (assess_id,))
    conn.commit()
    conn.close()
    return {"message": "Assessment deleted"}

# Change Log
@app.get("/api/changelog")
def list_change_log():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM change_log ORDER BY id DESC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@app.post("/api/changelog")
def create_change_log(data: ChangeLogCreate):
    conn = get_db()
    cur = conn.cursor()
    today = datetime.date.today()
    cur.execute(
        "INSERT INTO change_log (log_date, log_year, action_type, item_id, description, author) VALUES (?, ?, ?, ?, ?, ?)",
        (today.strftime('%Y-%m-%d'), today.year, data.action_type, data.item_id, data.description, data.author)
    )
    conn.commit()
    conn.close()
    return {"message": "Log recorded"}

# Excel Export
@app.get("/api/export-excel")
def export_excel():
    conn = get_db()
    cur = conn.cursor()

    wb = openpyxl.Workbook()
    # remove default sheet
    default_sheet = wb.active

    # Header styling
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    border_thin = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    # 1. Sheet: Laws
    ws_laws = wb.create_sheet(title="Laws")
    law_headers = ['Law_ID', 'ชื่อกฎหมาย', 'ประเภท', 'หน่วยงาน', 'วันที่มีผลบังคับ', 'สถานะ', 'ปีที่เพิ่มเข้าระบบ', 'ลิงก์ต้นฉบับ']
    ws_laws.append(law_headers)
    cur.execute("SELECT law_id, title, law_type, ministry, effective_date, status, year_added, source_link FROM laws ORDER BY law_id")
    for row in cur.fetchall():
        ws_laws.append(list(row))

    # 2. Sheet: Requirements
    ws_req = wb.create_sheet(title="Requirements")
    req_headers = ['Req_ID', 'Law_ID', 'ข้อ/มาตรา', 'สรุปข้อกำหนด', 'ฝ่ายรับผิดชอบ']
    ws_req.append(req_headers)
    cur.execute("SELECT req_id, law_id, section, summary, department FROM requirements ORDER BY req_id")
    for row in cur.fetchall():
        ws_req.append(list(row))

    # 3. Sheet: Assessments
    ws_assess = wb.create_sheet(title="Assessments")
    assess_headers = ['Assess_ID', 'Req_ID', 'วันที่ประเมิน', 'ผู้ประเมิน', 'ผล', 'หลักฐาน', 'แผนแก้ไข', 'กำหนดเสร็จ']
    ws_assess.append(assess_headers)
    cur.execute("SELECT assess_id, req_id, assessment_date, assessor, result, evidence, action_plan, due_date FROM assessments ORDER BY assess_id")
    for row in cur.fetchall():
        ws_assess.append(list(row))

    # 4. Sheet: Change_Log
    ws_log = wb.create_sheet(title="Change_Log")
    log_headers = ['วันที่', 'ปี', 'การเปลี่ยนแปลง', 'รายการ (Law_ID/Req_ID)', 'รายละเอียด', 'ผู้บันทึก']
    ws_log.append(log_headers)
    cur.execute("SELECT log_date, log_year, action_type, item_id, description, author FROM change_log ORDER BY id DESC")
    for row in cur.fetchall():
        ws_log.append(list(row))

    # Format sheets
    for ws in [ws_laws, ws_req, ws_assess, ws_log]:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.remove(default_sheet)
    conn.close()

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    filename = f"Legal_Register_Export_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# Re-seed API
@app.post("/api/reseed")
def reseed_database():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM change_log")
    cur.execute("DELETE FROM assessments")
    cur.execute("DELETE FROM requirements")
    cur.execute("DELETE FROM laws")
    conn.commit()
    seed_from_excel(conn)
    conn.close()
    return {"message": "Database successfully reloaded from Excel"}

# Serve frontend HTML
@app.get("/", response_class=HTMLResponse)
def index_page():
    html_file = os.path.join(BASE_DIR, "templates", "index.html")
    if os.path.exists(html_file):
        with open(html_file, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Legal Register System</h1><p>Template index.html not found.</p>"

if __name__ == "__main__":
    import uvicorn
    import webbrowser
    print("=======================================================")
    print("🚀 Legal Register & Compliance System is starting...")
    print("📍 URL: http://localhost:8000")
    print("=======================================================")
    webbrowser.open("http://localhost:8000")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
