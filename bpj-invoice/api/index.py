import os,json,io,zipfile,datetime,re,glob,urllib.parse
from http.server import BaseHTTPRequestHandler
from decimal import Decimal
from openpyxl import load_workbook
from supabase import create_client
ROOT=os.path.dirname(os.path.dirname(__file__))
TEMPLATES={
 'ss-front-pocket':'INVOICE  SS FRON  PCK.xlsx',
 'long-pen-pocket':'INVOICE (long+pen_pocket).xlsx',
 'short-pen-pocket':'INVOICE (short+pen_pocket).xlsx',
 'ls-back-logo':'INVOICE LS BACK LOGO.xlsx',
 'ss-no-pocket':'INVOICE SS NO PCK.xlsx'}
def db():return create_client(os.environ['SUPABASE_URL'],os.environ['SUPABASE_SERVICE_ROLE_KEY'])
def payload(row):
 return {'id':row['id'],'no':f"{row['invoice_no']:04d}",'date':row['invoice_date'],'template':row['template'],'customer':row['customer'],'items':row['items'],'total':float(row['total'])}
def excel(row):
 path=os.path.join(ROOT,'templates',TEMPLATES[row['template']]);wb=load_workbook(path);s=wb.active
 # Existing template images, positions, formatting and text remain untouched.
 s['G2']='NO:'+f"{row['invoice_no']:04d}"
 s['H3']=datetime.date.fromisoformat(row['invoice_date']);s['H3'].number_format='dd/mm/yyyy'
 s['A9']=row['customer']
 start=13 if row['template']=='ss-no-pocket' else 15
 for i,it in enumerate(row['items']):
  r=start+i
  s[f'A{r}']=it['item'];s[f'B{r}']=it['description'];s[f'F{r}']=it['qty'];s[f'G{r}']=it['price']
  s[f'H{r}']=f'=F{r}*G{r}'
  for col in 'GH':s[f'{col}{r}'].number_format='#,##0.00'
 # Fix original inconsistent total range in SS NO PCK, without changing appearance.
 if row['template']=='ss-no-pocket':
  s['H30']=f'=SUM(H{start}:H29)'
 s.print_options.horizontalCentered=True
 s.sheet_properties.pageSetUpPr.fitToPage=True;s.page_setup.fitToWidth=1;s.page_setup.fitToHeight=1
 s.page_setup.orientation='portrait';s.page_setup.paperSize=s.PAPERSIZE_A4
 s.print_area='A1:K52'
 b=io.BytesIO();wb.save(b);return b.getvalue()
def send(handler,status,obj=None,blob=None,mime='application/json',name=None):
 data=blob if blob is not None else json.dumps(obj,ensure_ascii=False,default=str).encode()
 handler.send_response(status);handler.send_header('Content-Type',mime);handler.send_header('Content-Length',str(len(data)))
 handler.send_header('Cache-Control','no-store');handler.send_header('X-Content-Type-Options','nosniff')
 if name:handler.send_header('Content-Disposition',f'attachment; filename="{name}"')
 handler.end_headers();handler.wfile.write(data)
class handler(BaseHTTPRequestHandler):
 def auth(self):
  token=os.environ.get('ADMIN_PASSWORD','')
  return bool(token) and self.headers.get('X-Admin-Password','')==token
 def do_GET(self):self.route('GET')
 def do_POST(self):self.route('POST')
 def do_DELETE(self):self.route('DELETE')
 def route(self,method):
  if not self.auth():return send(self,401,{'error':'Wrong admin password'})
  try:
   u=urllib.parse.urlparse(self.path);q=urllib.parse.parse_qs(u.query);path=u.path.removeprefix('/api')
   client=db()
   if path in ('','/','/invoices') and method=='GET':
    rows=client.table('invoices').select('*').order('invoice_no',desc=True).limit(1000).execute().data
    return send(self,200,{'invoices':[payload(x) for x in rows]})
   if path=='/next' and method=='GET':
    r=client.table('invoice_counter').select('last_no').eq('id',1).single().execute().data
    return send(self,200,{'next':f"{r['last_no']+1:04d}"})
   if path=='/invoices' and method=='POST':
    size=int(self.headers.get('Content-Length','0'))
    if size>100000:return send(self,413,{'error':'Request too large'})
    v=json.loads(self.rfile.read(size));date=datetime.date.fromisoformat(v['date'])
    template=v['template'];customer=str(v['customer']).strip()
    if template not in TEMPLATES or not customer or len(customer)>400:return send(self,400,{'error':'Invalid template or customer'})
    items=v['items'];limit=17 if template=='ss-no-pocket' else 15
    if not isinstance(items,list) or not 1<=len(items)<=limit:return send(self,400,{'error':f'Use 1–{limit} items'})
    clean=[];total=Decimal('0')
    for it in items:
     item=str(it.get('item','')).strip()[:70];desc=str(it.get('description','')).strip()[:200]
     qty=Decimal(str(it['qty']));price=Decimal(str(it['price']))
     if not item or not qty>0 or qty>1000000 or price<0 or price>1000000 or qty.as_tuple().exponent < -2 or price.as_tuple().exponent < -2:raise ValueError('Invalid item, quantity or price')
     clean.append({'item':item,'description':desc,'qty':float(qty),'price':float(price)})
     total+=qty*price
    if total>Decimal('9999999999'):raise ValueError('Total too large')
    r=client.rpc('create_invoice',{'p_date':str(date),'p_template':template,'p_customer':customer,'p_items':clean,'p_total':str(total)}).execute().data
    if isinstance(r,list):r=r[0]
    return send(self,201,payload(r))
   if path=='/export' and method=='GET':
    no=int(q['no'][0]);fmt=q.get('format',['xlsx'])[0]
    rows=client.table('invoices').select('*').eq('invoice_no',no).limit(1).execute().data
    if not rows:return send(self,404,{'error':'Not found'})
    if fmt=='xlsx':return send(self,200,blob=excel(rows[0]),mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',name=f'Invoice_{no:04d}.xlsx')
    return send(self,400,{'error':'PDF requires configured faithful Excel-to-PDF conversion service; see README'})
   if path=='/backup' and method=='GET':
    month=q.get('month',[''])[0]
    if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month):return send(self,400,{'error':'Invalid month'})
    rows=client.table('invoices').select('*').gte('invoice_date',month+'-01').lt('invoice_date',str((datetime.date.fromisoformat(month+'-01').replace(day=28)+datetime.timedelta(days=4)).replace(day=1))).order('invoice_no').limit(1000).execute().data
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
     for row in rows:z.writestr(f"Excel/Invoice_{row['invoice_no']:04d}.xlsx",excel(row))
     z.writestr('manifest.json',json.dumps([payload(x) for x in rows],indent=2))
    return send(self,200,blob=buf.getvalue(),mime='application/zip',name=f'Invoices_{month}.zip')
   return send(self,404,{'error':'Unknown endpoint'})
  except Exception as e:
   if isinstance(e,(ValueError,KeyError,ArithmeticError)):return send(self,400,{'error':str(e)})
   print('Server error:',repr(e));return send(self,500,{'error':'Server error; check Vercel logs'})
