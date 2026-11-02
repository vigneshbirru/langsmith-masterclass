import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

TODAY = dt.date.today()

DATA = [
 (1,"P1","Digitals AI Inc.","Generative AI Engineer – Fresher","Full-time","0–1 yr","GenAI, RAG, LangChain, Python","https://www.linkedin.com/jobs/search/?keywords=Generative%20AI%20Engineer%20Fresher&location=Hyderabad","LinkedIn","Verify before applying","","Startup. Strong GenAI match — confirm fresher eligibility before applying.",2),
 (2,"P1","RingTrunk","Agentic AI Intern","Internship (6 mo)","0 yr","LangGraph, Agentic AI, Python","https://www.linkedin.com/jobs/search/?keywords=Agentic%20AI%20Intern&location=Hyderabad","LinkedIn","Verify before applying","","Direct match to LangGraph / Agentic AI skills. Ask about PPO conversion.",2),
 (3,"P1","AgentAnalytics.AI","ML/AI Intern","Internship","0 yr","Machine Learning, Python, RAG","https://www.linkedin.com/jobs/search/?keywords=AgentAnalytics%20AI%20Intern&location=Hyderabad","LinkedIn","Verify before applying","","Analytics startup. Clarify stipend, duration, full-time conversion.",2),
 (4,"P1","Pragma Edge","AI Developer Trainee","Full-time (Trainee)","0–1 yr","Python, ML, LLM integration","https://www.pragmaedge.com/careers/","Company careers page","Verify before applying","","Hyderabad-based. Structured trainee programme.",2),
 (5,"P1","Tachyon Technologies","Python + AI/ML Graduate Intern","Internship","0 yr","Python, Machine Learning","https://www.linkedin.com/jobs/search/?keywords=Tachyon%20Technologies%20AI%20ML&location=Hyderabad","LinkedIn","Verify before applying","","Graduate internship. Confirm stipend and location.",2),
 (6,"P1","Flyo.ai","AI Engineering Intern","Internship","0 yr","GenAI, RAG, LangChain","https://www.linkedin.com/jobs/search/?keywords=Flyo.ai%20AI%20Engineer&location=Hyderabad","LinkedIn","Verify before applying","","AI product startup. Strong RAG portfolio helps.",2),
 (7,"P2","Tech Mahindra","AI/ML Trainee","Full-time","0–1 yr","ML, Python","https://careers.techmahindra.com","Company careers page","Verify before applying","","Filter Hyderabad on careers portal. Watch off-campus drives.",4),
 (8,"P2","Infosys","AI Engineer – Fresher","Full-time","0–1 yr","ML, GenAI","https://career.infosys.com","Company careers page","Verify before applying","","Apply on portal + register for off-campus drives.",4),
 (9,"P2","Wipro","AI/ML Associate","Full-time","0–1 yr","ML, Python","https://careers.wipro.com","Company careers page","Verify before applying","","Look at ELITE / Turbo fresher tracks.",4),
 (10,"P2","Accenture","Applied AI Associate","Full-time","0–1 yr","ML, GenAI","https://www.accenture.com/in-en/careers/jobsearch","Company careers page","Verify before applying","","Search 'AI Associate' + location Hyderabad.",4),
 (11,"P2","Cognizant","GenAI Engineer Trainee","Full-time","0–1 yr","GenAI, LLM, Python","https://careers.cognizant.com/global/en","Company careers page","Verify before applying","","GenAI upskilling tracks open periodically. Set job alert.",4),
 (12,"P2","TCS","AI/ML Developer – Fresher","Full-time","0–1 yr","ML, Python","https://ibegin.tcs.com/iBegin/jobs","Company careers page","Verify before applying","","TCS NQT / Smart Hiring is main route. Register early.",4),
 (13,"P2","Capgemini","Data Science Trainee","Full-time","0–1 yr","ML, Python","https://www.capgemini.com/in-en/careers/","Company careers page","Verify before applying","","Also search 'AI Engineer' — titles vary.",4),
 (14,"P2","Quantiphi","ML Engineer – Fresher","Full-time","0–1 yr","ML, GenAI, RAG","https://quantiphi.com/careers/","Company careers page","Verify before applying","","Strong GenAI practice — RAG/LangChain projects relevant.",4),
 (15,"P2","Tiger Analytics","Analyst – AI/ML","Full-time","0–1 yr","ML, Python","https://www.tigeranalytics.com/careers/","Company careers page","Verify before applying","","Case-study style interviews common.",4),
 (16,"P3","Fractal Analytics","AI Engineer – Fresher","Full-time","0–1 yr","ML, GenAI","https://fractal.ai/careers/","Company careers page","Verify before applying","","Check 'Imagineer' fresher programme.",6),
 (17,"P3","ValueLabs","AI/ML Engineer Trainee","Full-time","0–1 yr","ML, Python, GenAI","https://www.valuelabs.com/careers/","Company careers page","Verify before applying","","Hyderabad HQ. Hires freshers into AI pods.",6),
 (18,"P3","Cyient","AI/ML Engineer Trainee","Full-time","0–1 yr","ML, Python","https://careers.cyient.com","Company careers page","Verify before applying","","Engineering services, strong Hyderabad presence.",6),
 (19,"P3","Innominds","AI/ML Fresher","Full-time","0–1 yr","ML, GenAI","https://www.innominds.com/careers","Company careers page","Verify before applying","","Check fresher hiring status on portal.",6),
 (20,"P3","Darwinbox","AI Engineer (Fresher)","Full-time","0–1 yr","GenAI, LLM, RAG","https://darwinbox.com/careers","Company careers page","Verify before applying","","HR-tech SaaS with LLM features.",6),
 (21,"P3","Zenoti","ML Engineer","Full-time","0–1 yr","ML, Python","https://www.zenoti.com/careers","Company careers page","Verify before applying","","SaaS unicorn. Confirm fresher openness.",6),
 (22,"P3","Thomson Reuters","AI Engineer – Early Career","Full-time","0–1 yr","GenAI, RAG","https://careers.thomsonreuters.com","Company careers page","Verify before applying","","Legal-tech AI. Filter Hyderabad.",6),
 (23,"P3","Providence India","Data Science Associate","Full-time","0–1 yr","ML, Python","https://www.providence-india.in/careers","Company careers page","Verify before applying","","Healthcare AI. Check early-career openings.",6),
 (24,"P3","NxtWave","AI/ML Engineer","Full-time","0–1 yr","ML, GenAI","https://www.ccbp.in/careers","Company careers page","Verify before applying","","Ed-tech. Coding round first.",6),
 (25,"P4","Microsoft","AI/ML Intern","Internship","0 yr","ML, GenAI","https://careers.microsoft.com","Company careers page","Verify before applying","","Hyderabad IDC. Campus + off-campus.",7),
 (26,"P4","Google","AI/ML Intern","Internship","0 yr","ML, GenAI","https://careers.google.com","Company careers page","Verify before applying","","STEP / Summer internship. Short windows.",7),
 (27,"P4","Amazon","ML Intern","Internship","0 yr","Machine Learning","https://www.amazon.jobs","Company careers page","Verify before applying","","Hyd development centre. Also 'Applied Scientist Intern'.",7),
 (28,"P4","Adobe","AI/ML Intern","Internship","0 yr","ML, GenAI","https://careers.adobe.com","Company careers page","Verify before applying","","Check university/student section.",7),
 (29,"P4","Salesforce","AI Associate","Full-time","0–1 yr","GenAI, Python","https://careers.salesforce.com","Company careers page","Verify before applying","","Einstein AI team. Look under 'Early Career'.",7),
 (30,"P4","ServiceNow","AI/ML Intern","Internship","0 yr","ML, GenAI","https://careers.servicenow.com","Company careers page","Verify before applying","","Hyderabad office. Internships posted seasonally.",7),
]

HEADERS = ["S.No","Priority","Company","Role","Job Type","Experience","Skill Match","Apply Link","Source","Vacancy Status","HR Contact","Date Added","Follow-up Date","Application Status","Interview Stage","Outcome","Notes"]
WIDTHS  = [6,9,22,30,17,12,28,46,20,22,18,12,13,20,18,13,55]

HF = PatternFill("solid", fgColor="1F3864")
HFONT = Font(bold=True, color="FFFFFF", size=11)
BAND = PatternFill("solid", fgColor="F2F5FB")
P1F = PatternFill("solid", fgColor="FCE4D6")
LINK = Font(color="0563C1", underline="single", size=10)
BASE = Font(size=10)
THIN = Side(style="thin", color="BFBFBF")
BRD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)
CEN = Alignment(horizontal="center", vertical="top", wrap_text=True)

wb = Workbook()
ws = wb.active; ws.title = "Leads"
ws.merge_cells("A1:Q1")
ws["A1"] = "Hyderabad (HYD) AI/ML & Agentic AI — Fresher Job Tracker (30 leads)"
ws["A1"].font = Font(bold=True, size=14, color="1F3864")
ws.row_dimensions[1].height = 24
ws.merge_cells("A2:Q2")
ws["A2"] = f"Profile: B.E. | Python, ML, GenAI, RAG, LangChain, LangGraph  •  Generated: {TODAY:%d-%b-%Y}  •  Verify every vacancy before applying"
ws["A2"].font = Font(size=9, italic=True, color="595959")

HR_ROW = 4
for c,h in enumerate(HEADERS,1):
    cell = ws.cell(row=HR_ROW, column=c, value=h)
    cell.fill = HF; cell.font = HFONT; cell.border = BRD
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[HR_ROW].height = 32

for i,row in enumerate(DATA):
    r = HR_ROW+1+i
    sno,prio,comp,role,jtype,exp,skills,link,source,status,contact,notes,off = row
    vals = [sno,prio,comp,role,jtype,exp,skills,link,source,status,contact,TODAY,TODAY+dt.timedelta(days=off),"Not Started","—","Pending",notes]
    for c,v in enumerate(vals,1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = BASE; cell.border = BRD
        cell.alignment = CEN if c in (1,2,6,12,13,14,15,16) else WRAP
        if i % 2 == 1: cell.fill = BAND
        if c in (12,13): cell.number_format = "DD-MMM-YYYY"
    lc = ws.cell(row=r, column=8); lc.hyperlink = link; lc.font = LINK
    pc = ws.cell(row=r, column=2); pc.font = Font(bold=True, size=10)
    if prio == "P1": pc.fill = P1F
    elif prio == "P2": pc.fill = PatternFill("solid", fgColor="FFF2CC")

for c,w in enumerate(WIDTHS,1):
    ws.column_dimensions[get_column_letter(c)].width = w

last = HR_ROW + len(DATA)
ws.auto_filter.ref = f"A{HR_ROW}:Q{last}"
ws.freeze_panes = "D5"

dv1 = DataValidation(type="list", formula1='"Not Started,Applied,Follow-up Sent,Screening,Interview,Offer,Rejected,On Hold"', allow_blank=True)
dv2 = DataValidation(type="list", formula1='"—,HR Round,Technical Round,Managerial Round,Final Round"', allow_blank=True)
dv3 = DataValidation(type="list", formula1='"Pending,Selected,Rejected,Withdrawn,No Response"', allow_blank=True)
for dv in (dv1,dv2,dv3): ws.add_data_validation(dv)
dv1.add(f"N{HR_ROW+1}:N{last+60}"); dv2.add(f"O{HR_ROW+1}:O{last+60}"); dv3.add(f"P{HR_ROW+1}:P{last+60}")

rng = f"N{HR_ROW+1}:N{last+60}"
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Offer"'], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Rejected"'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006")))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Interview"'], fill=PatternFill("solid", fgColor="BDD7EE"), font=Font(color="1F3864", bold=True)))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Applied"'], fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(color="9C6500")))

# Dashboard
db = wb.create_sheet("Dashboard")
for col,w in {"A":4,"B":34,"C":14,"D":46}.items(): db.column_dimensions[col].width = w
db["B2"] = "Application Dashboard"; db["B2"].font = Font(bold=True, size=15, color="1F3864")
for c,h in (("B4","Metric"),("C4","Count"),("D4","Notes")):
    db[c] = h; db[c].fill = HF; db[c].font = HFONT; db[c].border = BRD
    db[c].alignment = Alignment(horizontal="center")
L = f"Leads!N{HR_ROW+1}:N{last}"
metrics = [
 ("Total leads", f"=COUNTA(Leads!C{HR_ROW+1}:C{last})","All 30 tracked leads"),
 ("P1 priority leads", f'=COUNTIF(Leads!B{HR_ROW+1}:B{last},"P1")',"Apply to these first"),
 ("Applications submitted", f'=COUNTA({L})-COUNTIF({L},"Not Started")',"Any status past 'Not Started'"),
 ("Not started", f'=COUNTIF({L},"Not Started")',"Still to apply"),
 ("In interview", f'=COUNTIF({L},"Interview")',"Active loops"),
 ("Offers", f'=COUNTIF({L},"Offer")',"🎉"),
 ("Rejections", f'=COUNTIF({L},"Rejected")',"Learn and move on"),
 ("Follow-ups due today", f'=COUNTIF(Leads!M{HR_ROW+1}:M{last},TODAY())',"Check Follow-up Date column"),
 ("Response rate", '=IFERROR((C9+C10)/MAX(C7,1),0)',"Offers + interviews ÷ submitted"),
]
for i,(lbl,f,note) in enumerate(metrics):
    r = 5+i
    db.cell(row=r,column=2,value=lbl).font = Font(size=10, bold=(i==0))
    c = db.cell(row=r,column=3,value=f); c.font = Font(size=11,bold=True); c.alignment = Alignment(horizontal="center")
    db.cell(row=r,column=4,value=note).font = Font(size=9, italic=True, color="595959")
    for col in (2,3,4):
        db.cell(row=r,column=col).border = BRD
        if i%2==1: db.cell(row=r,column=col).fill = BAND
db["C13"].number_format = "0.0%"

db["B16"] = "Status breakdown"; db["B16"].font = Font(bold=True, size=12, color="1F3864")
db["B17"] = "Status"; db["C17"] = "Count"
for c in ("B17","C17"):
    db[c].fill = HF; db[c].font = HFONT; db[c].border = BRD; db[c].alignment = Alignment(horizontal="center")
for i,s in enumerate(["Not Started","Applied","Follow-up Sent","Screening","Interview","Offer","Rejected","On Hold"]):
    r = 18+i
    db.cell(row=r,column=2,value=s).font = BASE
    cc = db.cell(row=r,column=3,value=f'=COUNTIF({L},"{s}")'); cc.font = BASE; cc.alignment = Alignment(horizontal="center")
    for col in (2,3): db.cell(row=r,column=col).border = BRD

# Priority Shortlist
ps = wb.create_sheet("Priority Shortlist")
for col,w in {"A":5,"B":24,"C":32,"D":34,"E":52}.items(): ps.column_dimensions[col].width = w
ps.merge_cells("A1:E1"); ps["A1"] = "Top 6 leads — apply to these first"
ps["A1"].font = Font(bold=True, size=14, color="1F3864")
for c,h in enumerate(["#","Company","Role","Why it fits your profile","Action"],1):
    cell = ps.cell(row=3,column=c,value=h); cell.fill = HF; cell.font = HFONT; cell.border = BRD
    cell.alignment = Alignment(horizontal="center", vertical="center")
top6 = [
 ("1","Digitals AI Inc.","Generative AI Engineer – Fresher","Direct GenAI + RAG + LangChain stack match","Open JD → confirm fresher accepted → apply with RAG project link"),
 ("2","RingTrunk","Agentic AI Intern","LangGraph / agentic orchestration is your differentiator","Ask about PPO conversion in first call"),
 ("3","AgentAnalytics.AI","ML/AI Intern","ML fundamentals + Python; startup ownership","Confirm stipend, duration, Hyderabad location"),
 ("4","Pragma Edge","AI Developer Trainee","Structured trainee track for 0–1 yr candidates","Check careers page for current AI openings"),
 ("5","Tachyon Technologies","Python + AI/ML Graduate Intern","Pure Python + ML graduate pipeline","Verify stipend and duration before applying"),
 ("6","Flyo.ai","AI Engineering Intern","GenAI product work — RAG portfolio decisive","Send 1-page project summary with application"),
]
for i,(n,co,role,why,act) in enumerate(top6):
    r = 4+i
    for c,v in enumerate([n,co,role,why,act],1):
        cell = ps.cell(row=r,column=c,value=v); cell.font = BASE; cell.border = BRD; cell.alignment = WRAP
        if i%2==1: cell.fill = BAND
    ps.row_dimensions[r].height = 34

# How to Use
hu = wb.create_sheet("How to Use")
hu.column_dimensions["A"].width = 4; hu.column_dimensions["B"].width = 110
hu["B2"] = "How to use this tracker"; hu["B2"].font = Font(bold=True, size=15, color="1F3864")
steps = [
 ("1. Verify before you apply","Every row reads 'Verify before applying'. Open the link, confirm the role is live and freshers (0–1 yr) are eligible."),
 ("2. Work top-down by priority","P1 (rows 1–6) are the closest matches to your Python / ML / GenAI / RAG / LangChain / LangGraph profile."),
 ("3. Update Application Status","Column N is a dropdown: Not Started → Applied → Follow-up Sent → Screening → Interview → Offer / Rejected. Dashboard refreshes automatically."),
 ("4. Track interview stages","Column O records the round; Column P records the outcome."),
 ("5. Respect the follow-up date","Column M auto-sets a follow-up date (P1 +2, P2 +4, P3 +6, P4 +7 days). Send a short polite follow-up on/after that date."),
 ("6. HR contact column","Left blank unless a verified public number/email exists. Use LinkedIn 'People' search for recruiters."),
 ("7. Portfolio tips","For GenAI/Agentic roles lead with: a RAG pipeline (LangChain + vector DB), a LangGraph multi-agent workflow, and one end-to-end ML project with metrics."),
 ("8. Honest disclaimer","Research/tracking template, not a guaranteed vacancy list. Some entries are internship-directory leads; some junior roles may still ask for prior experience."),
]
r = 4
for t,b in steps:
    hu.cell(row=r,column=2,value=t).font = Font(bold=True, size=11, color="1F3864"); r += 1
    c = hu.cell(row=r,column=2,value=b); c.font = Font(size=10); c.alignment = WRAP
    hu.row_dimensions[r].height = 44; r += 2

OUT = "Hyderabad_AI_ML_Fresher_Tracker.xlsx"
wb.save(OUT)
print("✅ Saved:", OUT)

# Colab: auto-download
try:
    from google.colab import files
    files.download(OUT)
except ImportError:
    pass