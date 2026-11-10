"""Black Ink PC V1 starter: local transaction register and manual coding.
Run: python black_ink.py  (Python 3.11+; pip install openpyxl)
No network access and no automatic journal posting.
"""
import csv
import datetime as dt
import hashlib
import os
import sqlite3
from collections import Counter
from pdf_import import parse_pdf
from accounting_engine import opening_from_workbook, build_preview, compute_balances, diagnostics, MONTHS
from approved_reference import approved_journal
from reporting import groups_from_reference, grouped_trial_balance, make_financials, round_financial_statement_rows, export_pdf, export_pdf_accountant_style
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from v14_features import (setup as setup_v14, is_approved, manual_lines, add_manual_journal,
                          month_controls, document_library, store_document)
from decimal import Decimal, InvalidOperation
from openpyxl import load_workbook
from report_exports import export_excel as export_report_excel, export_pdf as export_report_pdf
from ar_ap_tracking import setup as setup_arap, people as arap_people, set_transaction_party, open_tracker
from voice_coding import voice_window
from gd2go_reports import initialize as gd2go_setup, import_master as gd2go_import_master, year_end_window
from gd2go_journal import journal_entries
from operations_v14 import setup as setup_operations, open_operations, fifo_apply, ensure_operational_accounts

# Gd2Go has its own LOCAL application folder and a NEW, separate database.
# IMPORTANT: Never copy or connect to the Consulting or pilot database.
BASE_DATA = Path(os.environ.get('LOCALAPPDATA', Path.home()))
APPDIR = BASE_DATA / 'BlackInk_Gd2Go_V1_0'
APPDIR.mkdir(parents=True, exist_ok=True)
DB = APPDIR / 'gd2go.sqlite3'



def money(v):
    if v in (None, ''): return 0
    try: return int(Decimal(str(v).replace('$','').replace(',','').strip())*100)
    except (InvalidOperation, ValueError): return 0


def date_value(v):
    if isinstance(v, dt.datetime): return v.date().isoformat()
    if isinstance(v, dt.date): return v.isoformat()
    v = str(v or '').strip()
    for f in ('%Y-%m-%d','%m/%d/%Y','%d/%m/%Y','%b %d %Y'):
        try: return dt.datetime.strptime(v,f).date().isoformat()
        except ValueError: pass
    return ''


def signature(date, source, desc, debit, credit, srcfile, seq):
    # Source-row index differentiates identical purchases on one statement.
    s = '\x1f'.join(map(str,[date, source,desc,debit,credit,srcfile,seq]))
    return hashlib.sha256(s.encode('utf-8')).hexdigest()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Black Ink - Gd2Go | Version 1.4.8 — Home Subtitle Removed')
        self.geometry('1420x850')
        self.minsize(850,560)
        self._gl_editor = None
        self.configure(bg='#f4f7fa')
        self.cx = sqlite3.connect(DB)
        self.cx.execute('CREATE TABLE IF NOT EXISTS accounts(code TEXT PRIMARY KEY, description TEXT)')
        self.cx.execute('''CREATE TABLE IF NOT EXISTS transactions(
          id INTEGER PRIMARY KEY, fingerprint TEXT UNIQUE, date TEXT, transaction_date TEXT,
          source TEXT, description TEXT, debit INTEGER, credit INTEGER, user_gl TEXT,
          suggested_gl TEXT, source_file TEXT, card_kind TEXT DEFAULT 'Personal Visa')''')
        self.cx.commit()
        setup_v14(self.cx)
        setup_arap(self.cx)
        setup_operations(self.cx)
        self.data_folder = APPDIR
        gd2go_setup(self)
        ensure_operational_accounts(self.cx)
        self.month = tk.StringVar(value='2026-10')
        self.summary = tk.StringVar()
        self.build()
        self.refresh_months()
        self.refresh()

    def build(self):
        self.configure(bg='#f7f6f1')
        self.style = ttk.Style(self)
        try: self.style.theme_use('clam')
        except tk.TclError: pass
        self.style.configure('BI.TButton',font=('Segoe UI',10,'bold'),padding=(14,9))
        self.style.configure('BI.Treeview',rowheight=25,font=('Segoe UI',10))
        self.style.configure('BI.Treeview.Heading',font=('Segoe UI',10,'bold'))
        self.style.configure('BI.TNotebook.Tab',font=('Segoe UI',10,'bold'),padding=(16,8))
        # Main navigation is already provided by the navy Quick View bar.
        # Hide the redundant Home / Transaction Review tab strip, while retaining notebook page switching.
        self.style.layout('Hidden.TNotebook.Tab', [])
        header=tk.Frame(self,bg='#f7f6f1',height=118)
        header.pack(fill='x');header.pack_propagate(False)
        # Text and pen occupy distinct sides of the header; no overlap.
        brand=tk.Frame(header,bg='#f7f6f1');brand.pack(side='left',fill='y',padx=(35,0),pady=15)
        tk.Label(brand,text='BLACK INK - GD2GO',font=('Georgia',29,'bold'),bg='#f7f6f1',fg='#152d42').pack(anchor='w')
        logo=Path(__file__).with_name('black_ink_pen_header_sized.png')
        if logo.exists():
            self.pen_img=tk.PhotoImage(file=str(logo))
            tk.Label(header,image=self.pen_img,bg='#f7f6f1').pack(side='right',padx=(10,35))
        tk.Frame(self,bg='#bd9855',height=3).pack(fill='x')
        nav=tk.Frame(self,bg='#132e45');nav.pack(fill='x')
        tk.Label(nav,text='BLACK INK - GD2GO  |',bg='#132e45',fg='#e4ba70',font=('Segoe UI',9,'bold'),padx=14).pack(side='left')
        self.pages=ttk.Notebook(self,style='Hidden.TNotebook')
        self.home=tk.Frame(self.pages,bg='#f8f7f4')
        self.review=tk.Frame(self.pages,bg='#f4f7fa')
        self.pages.add(self.home,text='Home')
        self.pages.add(self.review,text='Transaction Review')
        for label,action in [('Home',lambda:self.pages.select(self.home)),
           ('Transactions',lambda:self.pages.select(self.review)),
           ('Journal Entries',lambda:year_end_window(self,'Adjusting JEs')),
           ('General Ledger',lambda:year_end_window(self,'General Ledger Detail')),
           ('Reports',self.financial_hub),
           ('Documents',self.document_hub),('AR / AP',lambda:open_tracker(self)),
           ('POS / Inventory',lambda:open_operations(self))]:
            tk.Button(nav,text=label,command=action,bg='#132e45',fg='white',
              activebackground='#244861',activeforeground='white',bd=0,
              font=('Segoe UI',10,'bold'),padx=16,pady=11).pack(side='left')
        self.pages.pack(fill='both',expand=True)
        self._build_home()
        review=self.review
        # Advanced actions remain available without overwhelming the coding workspace.
        tools=tk.Frame(review,bg='#e5edf4');tools.pack(fill='x')
        primary=tk.Frame(tools,bg='#e5edf4');primary.pack(fill='x',padx=8)
        tk.Label(primary,text='TRANSACTION REVIEW',font=('Segoe UI',12,'bold'),bg='#e5edf4',fg='#17374f').pack(side='left',padx=12)
        tk.Label(primary,text='Accounting month:',bg='#e5edf4',font=('Segoe UI',10,'bold')).pack(side='left',padx=(22,4))
        self.periods=ttk.Combobox(primary,textvariable=self.month,width=12,state='readonly')
        self.periods.pack(side='left',padx=3,pady=5)
        self.periods.bind('<<ComboboxSelected>>',lambda _:self.refresh())
        self.advanced_visible=tk.BooleanVar(value=False)
        def toggle_advanced():
            if self.advanced_visible.get():advanced.pack(fill='x',after=primary)
            else:advanced.pack_forget()
        # AR/AP and POS have dedicated top-level navigation. Do not duplicate
        # them on the transaction-review toolbar.
        ttk.Button(primary,text='Voice GL',command=lambda:voice_window(self)).pack(side='right',padx=(4,8))
        ttk.Checkbutton(primary,text='Advanced tools',variable=self.advanced_visible,command=toggle_advanced).pack(side='right',padx=(4,8))
        advanced=tk.Frame(tools,bg='#e5edf4')
        for label,fn in [('Import Gd2Go Master',lambda:gd2go_import_master(self)),('Import Workbook',self.import_workbook),('Import CSV',self.import_csv),
            ('Import Individual PDF',self.import_pdf),('Month Approval',lambda:month_controls(self)),
            ('Back Up Database',self.backup_db),('Export Month CSV',self.export_csv),
            ('Open Data Folder',lambda:os.startfile(str(APPDIR)) if os.name=='nt' else messagebox.showinfo('Data',str(APPDIR)))]:
            ttk.Button(advanced,text=label,command=fn).pack(side='left',padx=3,pady=6)
        tk.Label(review,textvariable=self.summary,bg='#f4f7fa',anchor='w',padx=18,font=('Segoe UI',10)).pack(fill='x',pady=8)
        pane=tk.PanedWindow(review,orient='horizontal',sashwidth=7,bg='#e5edf4');pane.pack(fill='both',expand=True,padx=12,pady=(0,8))
        self.review_pane=pane
        lf=tk.LabelFrame(pane,text='Chart of accounts • reference',font=('Segoe UI',10,'bold'))
        self.acct=ttk.Treeview(lf,columns=('code','name'),show='headings',height=22,style='BI.Treeview')
        for k,label,width in [('code','GL',58),('name','Description',224)]:
            self.acct.heading(k,text=label)
            self.acct.column(k,width=width,minwidth=55 if k=='code' else 145,stretch=(k=='name'),anchor='w')
        self.acct.pack(fill='both',expand=True,padx=5,pady=5)
        pane.add(lf,minsize=235,width=300)
        rf=tk.Frame(pane,bg='#f4f7fa')
        transaction_tools=tk.Frame(rf,bg='#f4f7fa')
        transaction_tools.pack(fill='x',padx=3,pady=(1,4))
        tk.Label(transaction_tools,text='Double-click GL or Supplier / Customer to edit  •  Enter saves  •  Esc cancels',
                 bg='#f4f7fa',fg='#17374f',font=('Segoe UI',9,'bold')).pack(side='left',padx=4)
        payments=tk.Menubutton(transaction_tools,text='Payment allocation ▾',relief='raised',
                               font=('Segoe UI',9),padx=9,pady=3)
        menu=tk.Menu(payments,tearoff=0)
        menu.add_command(label='Apply to selected invoice (manual / partial)',command=self.apply_selected_to_arap)
        menu.add_command(label='Apply oldest invoice first (FIFO)',command=self.auto_apply_arap)
        payments.configure(menu=menu)
        payments.pack(side='right',padx=4)
        def toggle_chart():
            if lf.winfo_manager():
                pane.forget(lf)
                show_chart.configure(text='Show chart')
            else:
                pane.add(lf,before=rf,minsize=235,width=300)
                show_chart.configure(text='Hide chart')
        show_chart=ttk.Button(transaction_tools,text='Hide chart',command=toggle_chart)
        show_chart.pack(side='right',padx=4)
        cols=('date','source','description','user_gl','party','debit','credit','suggestion')
        transaction_area=ttk.Frame(rf)
        transaction_area.pack(fill='both',expand=True)
        # Keep both scrollbars in the same grid as the table: the horizontal
        # scrollbar stays visible even on smaller screens and at 125% scaling.
        transaction_area.grid_rowconfigure(0,weight=1)
        transaction_area.grid_columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(transaction_area,columns=cols,show='headings',style='BI.Treeview')
        for k,label,width in [('date','Date',80),('source','Source',100),
                              ('description','Transaction Description',235),
                              ('user_gl','USER GL',78),('party','Supplier / Customer',166),
                              ('debit','Withdrawal',100),('credit','Deposit',100),
                              ('suggestion','Suggested GL',95)]:
            self.tree.heading(k,text=label)
            self.tree.column(k,width=width,minwidth=width,stretch=False,
                             anchor='e' if k in ('debit','credit') else 'center' if k in ('user_gl','suggestion') else 'w')
        self.tree.tag_configure('coded',background='#edf8ef')
        self.tree.tag_configure('uncoded',background='#fff9e8')
        self.tree.tag_configure('statement_heading',background='#dbe7f2',foreground='#122c43',font=('Segoe UI',10,'bold'))
        self.tree.tag_configure('statement_gap',background='#ffffff')
        scroll=ttk.Scrollbar(transaction_area,orient='vertical',command=self.tree.yview)
        hscroll=ttk.Scrollbar(transaction_area,orient='horizontal',command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll.set,xscrollcommand=hscroll.set)
        self.tree.grid(row=0,column=0,sticky='nsew')
        scroll.grid(row=0,column=1,sticky='ns')
        hscroll.grid(row=1,column=0,sticky='ew')
        self.tree.bind('<Double-1>',self.code_selected)
        self.tree.bind('<Control-r>',lambda _:self.attach_selected_receipt())
        pane.add(rf,minsize=580)

    def _build_home(self):
        """Compact, scrollable dashboard: every action remains visible at any screen size."""
        # Keep the dashboard in its own scroll region rather than allowing a
        # third row of cards to disappear under the lower edge of the window.
        area=tk.Frame(self.home,bg='#f8f7f4')
        area.pack(fill='both',expand=True)
        canvas=tk.Canvas(area,bg='#f8f7f4',highlightthickness=0,bd=0)
        scroll=ttk.Scrollbar(area,orient='vertical',command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right',fill='y')
        canvas.pack(side='left',fill='both',expand=True)
        content=tk.Frame(canvas,bg='#f8f7f4')
        content_id=canvas.create_window((0,0),window=content,anchor='nw')
        def on_content_resize(_event=None):
            canvas.configure(scrollregion=canvas.bbox('all'))
        def on_canvas_resize(event):
            canvas.itemconfigure(content_id,width=event.width)
        content.bind('<Configure>',on_content_resize)
        canvas.bind('<Configure>',on_canvas_resize)
        # Scroll the home dashboard only while the pointer is over its area;
        # other tabs and popup windows keep their normal wheel behaviour.
        def wheel(event):
            if getattr(event,'delta',0):
                canvas.yview_scroll(-int(event.delta/120) if abs(event.delta)>=120 else (-1 if event.delta>0 else 1),'units')
            elif getattr(event,'num',None) in (4,5):
                canvas.yview_scroll(-1 if event.num==4 else 1,'units')
        def bind_wheel(_event):
            canvas.bind_all('<MouseWheel>',wheel)
            canvas.bind_all('<Button-4>',wheel)
            canvas.bind_all('<Button-5>',wheel)
        def unbind_wheel(_event):
            canvas.unbind_all('<MouseWheel>')
            canvas.unbind_all('<Button-4>')
            canvas.unbind_all('<Button-5>')
        area.bind('<Enter>',bind_wheel)
        area.bind('<Leave>',unbind_wheel)

        banner=tk.Frame(content,bg='#f8f7f4')
        banner.pack(fill='x',padx=36,pady=(18,5))
        tk.Label(banner,text='Welcome to Black Ink',bg='#f8f7f4',fg='#162e43',font=('Georgia',22)).pack(anchor='w')
        tk.Label(banner,text='Select your accounting period and choose where you want to work.',
          bg='#f8f7f4',fg='#596673',font=('Segoe UI',10)).pack(anchor='w',pady=(4,0))
        period=tk.Frame(content,bg='#f8f7f4')
        period.pack(fill='x',padx=36,pady=(12,12))
        tk.Label(period,text='ACCOUNTING PERIOD',bg='#f8f7f4',fg='#786b51',font=('Segoe UI',10,'bold')).pack(side='left',padx=(0,12))
        ttk.Combobox(period,textvariable=self.month,state='readonly',values=[f'{y}-{m:02d}' for y,m in [(2026,n) for n in range(8,13)]+[(2027,n) for n in range(1,8)]],width=12).pack(side='left')
        tk.Label(period,text='  Applies to transactions and documents',bg='#f8f7f4',fg='#66717d',font=('Segoe UI',9)).pack(side='left')
        grid=tk.Frame(content,bg='#f8f7f4')
        grid.pack(fill='x',padx=25,pady=(0,5))
        tiles=[('Document Library','Upload bank statements, credit cards and receipts',self.document_hub),
         ('Transaction Review','Code transactions using the chart of accounts',lambda:self.pages.select(self.review)),
         ('Journal Entries','Review or post balanced adjusting entries',lambda:year_end_window(self,'Adjusting JEs')),
         ('General Ledger','View account balances and journal entries',lambda:year_end_window(self,'General Ledger Detail')),
         ('Trial Balance','Comparative balances and account groupings',lambda:year_end_window(self,'Trial Balance')),
         ('Financial Statements','Internal statements and professional PDFs',self.financial_hub),
         ('AR / AP','Track supplier and customer invoices and aging',lambda:open_tracker(self)),
         ('POS / Inventory','Import POS sales, costs and reconcile inventory',lambda:open_operations(self))]
        for i,(label,detail,cmd) in enumerate(tiles):
            row,col=divmod(i,3)
            tile=tk.Frame(grid,bg='#17374f',highlightbackground='#bc9858',highlightthickness=1,height=112)
            tile.grid(row=row,column=col,padx=7,pady=7,sticky='nsew')
            tile.pack_propagate(False)
            tk.Button(tile,text=label+'  →',bg='#17374f',fg='#ffffff',activebackground='#244861',
              activeforeground='#ffffff',font=('Segoe UI',12,'bold'),relief='flat',anchor='w',
              command=cmd,padx=16,pady=11).pack(fill='x')
            tk.Label(tile,text=detail,bg='#17374f',fg='#d4dee5',font=('Segoe UI',9),
              wraplength=300,justify='left',padx=16,pady=2,anchor='nw').pack(fill='both',expand=True)
        for i in range(3): grid.columnconfigure(i,weight=1,uniform='home_cards')
        tk.Label(content,text='BLACK INK - GD2GO  •  VERSION 1.4.8 HOME SUBTITLE REMOVED  •  SEPARATE DATABASE',
           fg='#a38755',bg='#f8f7f4',font=('Segoe UI',9,'bold')).pack(pady=(8,18))

    def _action_hub(self,title,subtitle,actions):
        win=tk.Toplevel(self)
        win.title('Black Ink — '+title)
        win.geometry('680x360')
        win.configure(bg='#f8f7f4')
        tk.Label(win,text=title,bg='#f8f7f4',fg='#17374f',font=('Georgia',23,'bold')).pack(anchor='w',padx=30,pady=(20,6))
        tk.Label(win,text=subtitle,bg='#f8f7f4',fg='#596673',font=('Segoe UI',11)).pack(anchor='w',padx=30,pady=(0,15))
        tabs=ttk.Notebook(win,style='BI.TNotebook')
        tabs.pack(fill='both',expand=True,padx=26,pady=(0,20))
        for label,action in actions:
            text=label.split('   ',1)[-1]
            panel=tk.Frame(tabs,bg='#f8f7f4')
            tabs.add(panel,text=text)
            tk.Label(panel,text=text,bg='#f8f7f4',fg='#17374f',font=('Segoe UI',15,'bold')).pack(anchor='w',padx=22,pady=(24,10))
            def execute(fn=action):
                win.destroy()
                self.after(20,fn)
            tk.Button(panel,text='Open '+text if not text.startswith('Import') else text,
                      command=execute,bg='#17374f',fg='white',activebackground='#244861',activeforeground='white',
                      relief='flat',padx=20,pady=12,font=('Segoe UI',12,'bold')).pack(anchor='w',padx=22)
        return win

    def document_hub(self):
        from v14_features import import_month_documents
        def import_all():
            month=self.month.get()
            try:
                info=import_month_documents(self,month)
                if info is None:return
                note=(f"{month}: {info['files']} documents reviewed; {info['new']} new transactions; "
                      f"{info['duplicates']} already recorded; {info['excluded']} outside month.")
                if info['issues']:note+='\nFiles needing attention: '+'; '.join(info['issues'][:5])
                messagebox.showinfo('Monthly import summary',note,parent=self)
                self.pages.select(self.review)
            except Exception as exc:messagebox.showerror('Document import',str(exc),parent=self)
        self._action_hub('Document Library','The selected accounting month applies to all three actions.',[
            ('1   Upload Documents',lambda:document_library(self,mode='upload')),
            ('2   Import Documents',import_all),
            ('3   Review Documents',lambda:document_library(self,mode='review'))])

    def journal_hub(self):
        year_end_window(self,'Adjusting JEs')

    def financial_hub(self):
        self._action_hub('Financial Statements','Choose internal statements or the existing professional PDF.',[
            ('1   Internal Financial Statements',lambda:year_end_window(self,'Financial Statements')),
            ('2   External Professional PDF',lambda:year_end_window(self,'Financial Statements',True))])

    def refresh_months(self):
        # Always show the full 2027 fiscal year, even when a month has no transactions.
        fiscal_months = [f'{year:04d}-{month:02d}' for year, month in
                         [(2026, m) for m in range(8, 13)] + [(2027, m) for m in range(1, 8)]]
        # Keep any previously imported dates accessible, including older fiscal years.
        found = [x[0] for x in self.cx.execute(
            'SELECT DISTINCT substr(date,1,7) FROM transactions ORDER BY 1') if x[0]]
        values = sorted(set(fiscal_months + found))
        self.periods['values'] = values
        if self.month.get() not in values:
            self.month.set('2026-09')

    def refresh(self):
        self._discard_gl_editor()
        for i in self.acct.get_children():self.acct.delete(i)
        for code,name in self.cx.execute('SELECT code,description FROM accounts ORDER BY code'):
            clean_name=str(name or '').strip()
            if clean_name.startswith(str(code)+' '):
                clean_name=clean_name[len(str(code)):].strip()
            self.acct.insert('', 'end', values=(code,clean_name))
        for i in self.tree.get_children():self.tree.delete(i)
        mo=self.month.get()
        # Present sources as five statement blocks, in a consistent order, not interleaved by date.
        # Only presentation changes here: transaction IDs, GL coding, and accounting calculations stay untouched.
        rows=list(self.cx.execute('SELECT t.id,t.date,t.source,t.description,t.debit,t.credit,t.user_gl,t.suggested_gl,p.name FROM transactions t LEFT JOIN counterparties p ON t.counterparty_id=p.id WHERE substr(t.date,1,7)=? ORDER BY CASE t.source WHEN "BMO Chequing" THEN 0 WHEN "BMO Savings" THEN 1 WHEN "TD Visa" THEN 2 WHEN "BMO Personal Chequing" THEN 3 WHEN "BMO Personal Mastercard" THEN 4 ELSE 5 END,t.date,t.id',(mo,)))
        labels = {
            'BMO Chequing': 'BLACK INK — BMO CHEQUING',
            'BMO Savings': 'BLACK INK — BMO SAVINGS',
            'TD Visa': 'TD PERSONAL VISA',
            'BMO Personal Chequing': 'BMO PERSONAL CHEQUING',
            'BMO Personal Mastercard': 'BMO PERSONAL MASTERCARD',
        }
        sources=['BMO Chequing','BMO Savings','TD Visa','BMO Personal Chequing','BMO Personal Mastercard']
        sources += sorted({r[2] for r in rows if r[2] not in sources})
        coded=0; deb=cred=0
        for group_index,src in enumerate(sources):
            group=[r for r in rows if r[2]==src]
            if not group:
                continue
            if group_index and self.tree.get_children():
                self.tree.insert('', 'end', iid=f'gap_{group_index}', values=('',)*8, tags=('statement_gap',))
            self.tree.insert('', 'end', iid=f'heading_{group_index}',
                             values=(mo, labels.get(src,src.upper()), f'{len(group)} transactions', '', '', '', '', ''),
                             tags=('statement_heading',))
            for id,d,source,desc,dr,cr,gl,sugg,party in group:
                coded+=bool(gl);deb+=dr;cred+=cr
                self.tree.insert('', 'end',iid=str(id),values=(d,source,desc,gl or '',party or '',format(dr/100,',.2f') if dr else '',format(cr/100,',.2f') if cr else '',sugg or ''),tags=('coded' if gl else 'uncoded',))
        self.summary.set(f'[{"APPROVED — LOCKED" if is_approved(self.cx,mo) else "OPEN"}]  {mo}  |  {len(rows)} transactions  |  {coded} coded  |  {len(rows)-coded} not coded (personal Visa blanks may be intentional)  |  Withdrawals/charges ${deb/100:,.2f}  |  Deposits/credits ${cred/100:,.2f}')

    def insert_row(self, data):
        date,transaction_date,src,desc,dr,cr,gl,source_file,seq=data
        if not date or not desc: return False
        f=signature(date,src,desc,dr,cr,source_file,seq)
        self.cx.execute('INSERT OR IGNORE INTO transactions(fingerprint,date,transaction_date,source,description,debit,credit,user_gl,source_file) VALUES(?,?,?,?,?,?,?,?,?)',(f,date,transaction_date,src,desc,dr,cr,gl,source_file))
        return self.cx.execute('SELECT changes()').fetchone()[0]>0

    def import_workbook(self):
        if is_approved(self.cx,self.month.get()):
            messagebox.showwarning('Month locked','Reopen this month before importing.');return
        path=filedialog.askopenfilename(filetypes=[('Excel workbook','*.xlsx')]);
        if not path:return
        try:
            wb=load_workbook(path,read_only=True,data_only=True)
            if 'Transaction Review' not in wb.sheetnames:raise ValueError('Transaction Review tab not found')
            count=0; accounts=0
            for i,original in enumerate(wb['Transaction Review'].iter_rows(values_only=True),1):
                r=list(original)+[None]*10
                if r[0] and str(r[0]).strip().isdigit() and r[1] and not r[2]:
                    self.cx.execute('INSERT OR REPLACE INTO accounts(code,description) VALUES(?,?)',(str(r[0]).strip(),str(r[1])));accounts+=1
                date=date_value(r[0])
                if date and r[2] and r[3] and (r[4] is not None or r[5] is not None):
                    gl=str(r[6]).strip().replace('.0','') if r[6] is not None else ''
                    count+=self.insert_row((date,date_value(r[1]),str(r[2]),str(r[3]),money(r[4]),money(r[5]),gl,str(r[9] or Path(path).name),i))
            self.cx.commit();store_document(self.cx,APPDIR,path,self.month.get(),'Workbooks','Excel workbook');self.refresh_months();self.refresh()
            messagebox.showinfo('Import',f'Imported {count} new transaction rows and {accounts} chart-of-accounts lines.\nOriginal Excel workbook was not changed.')
        except Exception as exc:messagebox.showerror('Import failed',str(exc))

    def import_csv(self):
        if is_approved(self.cx,self.month.get()):
            messagebox.showwarning('Month locked','Reopen this month before importing.');return
        path=filedialog.askopenfilename(filetypes=[('CSV files','*.csv')]);
        if not path:return
        try:
            with open(path,encoding='utf-8-sig',newline='') as f: rows=list(csv.DictReader(f))
            required={'Date','Source','Description','Withdrawal','Deposit'}
            if not rows or not required.issubset(rows[0]):raise ValueError('Required CSV headers: Date, Source, Description, Withdrawal, Deposit. Optional: TransactionDate, UserGL')
            n=0
            for i,r in enumerate(rows,1):
                n+=self.insert_row((date_value(r['Date']),date_value(r.get('TransactionDate')),r['Source'],r['Description'],money(r['Withdrawal']),money(r['Deposit']),str(r.get('UserGL','')).strip(),Path(path).name,i))
            self.cx.commit();store_document(self.cx,APPDIR,path,self.month.get(),'Statements','CSV');self.refresh_months();self.refresh();messagebox.showinfo('CSV import',f'{n} new rows imported. Re-importing the same file will not duplicate its rows.')
        except Exception as exc:messagebox.showerror('CSV import failed',str(exc))

    def backup_db(self, quiet=False):
        folder=APPDIR / 'Backups';folder.mkdir(exist_ok=True)
        name='blackink_' + dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.sqlite3'
        dest=folder/name
        with sqlite3.connect(dest) as copy:self.cx.backup(copy)
        if not quiet:messagebox.showinfo('Backup complete',f'Database copied to:\n{dest}\n\nYour accounting data remains in the original database.')
        return dest

    def import_pdf(self):
        if is_approved(self.cx,self.month.get()):
            messagebox.showwarning('Month locked','Reopen this month before importing.');return
        paths=filedialog.askopenfilenames(title='Choose BMO business/personal, BMO Mastercard or TD Visa PDFs',filetypes=[('PDF statements','*.pdf')])
        if not paths:return
        extracted=[];problems=[]
        for path in paths:
            try:
                kind,rows=parse_pdf(path)
                for r in rows:r['file']=Path(path).name
                extracted.extend(rows)
            except Exception as e:problems.append(f'{Path(path).name}: {e}')
        if not extracted:
            messagebox.showerror('Nothing imported','\n'.join(problems) or 'No transactions found')
            return
        mo=self.month.get()
        eligible=[r for r in extracted if r['date'][:7]==mo]
        if not eligible:
            messagebox.showwarning('No matching dates',f'No rows posted in {mo}.\nChoose the correct accounting month before importing.\n'+'\n'.join(problems))
            return
        existing=Counter(tuple(r) for r in self.cx.execute('SELECT date,source,description,debit,credit FROM transactions'))
        seen=Counter();new=[];duplicate=[]
        for row in eligible:
            key=(row['date'],row['source'],row['description'],row['debit'],row['credit'])
            seen[key]+=1
            (duplicate if seen[key]<=existing[key] else new).append(row)
        dlg=tk.Toplevel(self);dlg.title(f'PDF Import Preview — {mo}');dlg.geometry('1110x630');dlg.grab_set()
        hdr=f'{len(eligible)} transactions for {mo} | {len(new)} candidates | {len(duplicate)} possible duplicates | {len(extracted)-len(eligible)} other-month rows excluded'
        tk.Label(dlg,text=hdr,anchor='w',font=('Segoe UI',11,'bold')).pack(fill='x',padx=12,pady=10)
        tk.Label(dlg,text='Review merchant descriptions and amounts before import. Only new candidates are added. No GL account is assigned automatically.',anchor='w').pack(fill='x',padx=12)
        cols=('date','source','desc','debit','credit','file')
        table=ttk.Treeview(dlg,columns=cols,show='headings')
        for key,title,width in [('date','Posting Date',100),('source','Source',110),('desc','Description',425),('debit','Charge / Withdrawal',125),('credit','Credit / Deposit',125),('file','PDF',180)]:
            table.heading(key,text=title);table.column(key,width=width,anchor='w')
        for r in new:
            table.insert('', 'end',values=(r['date'],r['source'],r['description'],f"{r['debit']/100:,.2f}" if r['debit'] else '',f"{r['credit']/100:,.2f}" if r['credit'] else '',r['file']))
        table.pack(fill='both',expand=True,padx=12,pady=12)
        if problems:tk.Label(dlg,text='Unsupported files: '+' | '.join(problems),foreground='#9b3223',wraplength=1000).pack()
        def approve():
            if not new:
                for original in paths:
                    if Path(original).name not in [p.split(':',1)[0] for p in problems]:
                        store_document(self.cx,APPDIR,original,mo,'Statements','Bank / Visa PDF')
                dlg.destroy();messagebox.showinfo('No transaction changes','No new transactions imported. Original statements archived in the Document Library.');return
            try:
                backup=self.backup_db(quiet=True)
                count=0
                for seq,r in enumerate(new,1):
                    count+=self.insert_row((r['date'],r['transaction_date'],r['source'],r['description'],r['debit'],r['credit'],'',r['file'],seq))
                self.cx.commit()
                for original in paths:
                    # Capture every successfully parsed PDF in an unchanged archive.
                    if Path(original).name not in [p.split(':',1)[0] for p in problems]:
                        store_document(self.cx,APPDIR,original,mo,'Statements','Bank / Visa PDF')
                dlg.destroy();self.refresh_months();self.refresh()
                messagebox.showinfo('PDF import complete',f'{count} new transactions imported.\n{len(duplicate)} potential duplicates skipped.\n\nSafety backup saved:\n{backup}\n\nPlease reconcile against statements before coding.')
            except Exception as e:
                self.cx.rollback();messagebox.showerror('Import failed',str(e))
        bar=tk.Frame(dlg);bar.pack(pady=8)
        ttk.Button(bar,text=f'Import {len(new)} New Rows',command=approve).pack(side='left',padx=12)
        ttk.Button(bar,text='Cancel — Import Nothing',command=dlg.destroy).pack(side='left')

    def accounting_preview(self, initial_tab=None, export_prompt=False):
        # Unlike Consulting, Gd2Go uses an imported comparative year-end baseline.
        year_end_window(self, 'Balance Sheet' if initial_tab=='Financial Statements' else 'Trial Balance', export_prompt)

    def edit_transaction_party(self,iid):
        """Supplier/customer reference never changes GL account or journal entries."""
        from ar_ap_tracking import people
        p=people(self.cx)
        dlg=tk.Toplevel(self);dlg.title('Assign Supplier / Customer');dlg.geometry('520x260')
        desc=self.cx.execute('SELECT description FROM transactions WHERE id=?',(iid,)).fetchone()[0]
        ttk.Label(dlg,text='Assign party to transaction',font=('Segoe UI',14,'bold')).pack(anchor='w',padx=17,pady=(15,6))
        ttk.Label(dlg,text=desc,wraplength=440).pack(anchor='w',padx=17,pady=5)
        choices=['(None)']+[f'{name} [{kind}]' for _,name,kind in p]
        var=tk.StringVar(value='(None)')
        combo=ttk.Combobox(dlg,textvariable=var,values=choices,state='normal',width=52)
        combo.pack(anchor='w',padx=17,pady=9)
        ids={f'{name} [{kind}]':pid for pid,name,kind in p}
        # Typing filters the available suppliers/customers without adding
        # arbitrary names to the accounting register.
        def filter_parties(event=None):
            if event and event.keysym in ('Up','Down','Return','Escape','Tab'):
                return
            query=var.get().strip().casefold()
            combo['values']=[name for name in choices if query in name.casefold()]
        combo.bind('<KeyRelease>',filter_parties)
        ttk.Label(dlg,text='Type to search; select the registered supplier or customer.',
                  foreground='#5f7280').pack(anchor='w',padx=17)
        def save():
            if var.get() not in ids and var.get()!='(None)':
                messagebox.showwarning('Select a registered name',
                    'Choose a supplier/customer from the list, or create one in AR / AP Tracking.',parent=dlg)
                return
            set_transaction_party(self,int(iid),ids.get(var.get()))
            self.refresh();self.tree.selection_set(iid);self.tree.see(iid);dlg.destroy()
        ttk.Button(dlg,text='Save Supplier / Customer',command=save).pack(anchor='w',padx=17,pady=5)
        ttk.Label(dlg,text='Add new names in AR / AP Tracking.',foreground='#5f7280').pack(anchor='w',padx=17)

    def apply_selected_to_arap(self):
        s=self.tree.selection()
        if not s or not s[0].isdigit():
            messagebox.showwarning('Select transaction','Select a transaction first.');return
        open_tracker(self,int(s[0]))

    def auto_apply_arap(self):
        selection=self.tree.selection()
        if not selection or not selection[0].isdigit():
            messagebox.showwarning('Select a bank transaction','Select a bank transaction first.');return
        txid=int(selection[0])
        row=self.cx.execute('SELECT date,debit,credit,counterparty_id FROM transactions WHERE id=?',(txid,)).fetchone()
        if not row or not row[3]:
            messagebox.showwarning('Supplier / Customer','Assign a supplier or customer in the Party column first.');return
        kind='AR' if (row[2] or 0)>0 else 'AP' if (row[1] or 0)>0 else ''
        if not kind:return
        if not messagebox.askyesno('FIFO allocation preview',
            f'Allocate this {kind} bank transaction to the oldest outstanding invoice(s) for the assigned party?\n'
            'Allocation is informational only and will NOT post to the GL.'):
            return
        try:
            allocated,remaining=fifo_apply(self.cx,txid,kind)
            detail='\n'.join(f'{ref}: ${amount/100:,.2f}' for ref,amount in allocated) or 'No eligible invoices.'
            messagebox.showinfo('FIFO allocation completed',f'{detail}\nUnapplied balance: ${remaining/100:,.2f}\n\nReview in AR / AP Tracking.')
        except Exception as exc:messagebox.showerror('FIFO allocation blocked',str(exc))

    def _discard_gl_editor(self):
        editor=getattr(self,'_gl_editor',None)
        self._gl_editor=None
        if editor is not None:
            try: editor.destroy()
            except tk.TclError: pass

    def code_selected(self,event):
        # USER GL and independent Supplier / Customer column are editable.
        iid=self.tree.identify_row(event.y)
        column=self.tree.identify_column(event.x)
        if not iid or not iid.isdigit() or column not in ('#4','#5'):return
        if self.tree.identify_region(event.x,event.y)!='cell':return
        row=self.cx.execute('SELECT date,source,description,user_gl FROM transactions WHERE id=?',(iid,)).fetchone()
        if row is None:return
        if column=='#5':
            # Informational subledger metadata may be maintained independently of approved GL coding.
            self.edit_transaction_party(iid)
            return
        if is_approved(self.cx,row[0][:7]):
            messagebox.showwarning('Approved month','Reopen this month before changing transaction coding.')
            return
        self._discard_gl_editor()
        self.tree.see(iid)
        bbox=self.tree.bbox(iid,'#4')
        if not bbox:return
        x,y,w,h=bbox
        editor=ttk.Entry(self.tree,font=('Segoe UI',11))
        self._gl_editor=editor
        editor.insert(0,row[3] or '')
        editor.place(x=x,y=y,width=max(65,w),height=h)
        editor.focus_set()
        editor.selection_range(0,'end')
        def save(_=None):
            if self._gl_editor is not editor:return
            val=editor.get().strip()
            if val and not self.cx.execute('SELECT 1 FROM accounts WHERE code=?',(val,)).fetchone():
                messagebox.showwarning('Unknown GL account',f'{val} is not in the chart of accounts. Nothing was changed.')
                editor.focus_set();return
            try:
                with self.cx:
                    self.cx.execute('UPDATE transactions SET user_gl=? WHERE id=?',(val,iid))
            except sqlite3.Error as exc:
                messagebox.showerror('Could not save GL',str(exc));return
            self._discard_gl_editor()
            # Update the single row to preserve position and scrolling.
            data=list(self.tree.item(iid,'values'))
            data[3]=val
            self.tree.item(iid,values=data,tags=('coded' if val else 'uncoded',))
            self.tree.selection_set(iid)
            self._update_summary_only()
        editor.bind('<Return>',save)
        editor.bind('<KP_Enter>',save)
        editor.bind('<Escape>',lambda _:self._discard_gl_editor())

    def _update_summary_only(self):
        mo=self.month.get()
        rows=self.cx.execute("SELECT debit,credit,user_gl FROM transactions WHERE substr(date,1,7)=?",(mo,)).fetchall()
        coded=sum(bool(gl) for _,_,gl in rows)
        deb=sum(d or 0 for d,_,_ in rows);cred=sum(c or 0 for _,c,_ in rows)
        state='APPROVED — LOCKED' if is_approved(self.cx,mo) else 'OPEN'
        self.summary.set(f'[{state}]  {mo}  |  {len(rows)} transactions  |  {coded} coded  |  {len(rows)-coded} uncoded  |  Withdrawals/charges ${deb/100:,.2f}  |  Deposits/credits ${cred/100:,.2f}')

    def attach_selected_receipt(self):
        iid=self.tree.focus()
        if not iid or not iid.isdigit():
            messagebox.showinfo('Select transaction','Select a transaction first.')
            return
        row=self.cx.execute('SELECT date,source FROM transactions WHERE id=?',(iid,)).fetchone()
        if row is None:return
        path=filedialog.askopenfilename(parent=self,title='Attach receipt',filetypes=[('Receipt files','*.pdf *.png *.jpg *.jpeg'),('All files','*.*')])
        if path:
            try:
                store_document(self.cx,APPDIR,path,row[0][:7],'Receipts',row[1],int(iid))
                messagebox.showinfo('Receipt saved','Receipt archived and linked to this transaction.')
            except Exception as exc:messagebox.showerror('Receipt error',str(exc))

    def export_csv(self):
        path=filedialog.asksaveasfilename(defaultextension='.csv',initialfile=f'Black_Ink_{self.month.get()}_Review.csv',filetypes=[('CSV','*.csv')]);
        if not path:return
        with open(path,'w',encoding='utf-8-sig',newline='') as f:
            out=csv.writer(f);out.writerow(['Posting Date','Transaction Date','Source','Description','Withdrawal','Deposit','User GL','Suggested GL','Original Source File'])
            for r in self.cx.execute('SELECT date,transaction_date,source,description,debit,credit,user_gl,suggested_gl,source_file FROM transactions WHERE substr(date,1,7)=? ORDER BY date,id',(self.month.get(),)):
                out.writerow([*r[:4],f'{r[4]/100:.2f}',f'{r[5]/100:.2f}',*r[6:]])
        messagebox.showinfo('Export',f'Saved {path}')

if __name__=='__main__':
    from gd2go_auto_update import check_on_launch
    if not check_on_launch():
        App().mainloop()
