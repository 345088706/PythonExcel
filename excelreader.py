# encoding: utf-8
from docx import Document
from docx.shared import Inches
import sys,json,os,zipfile,xlrd,hashlib
reload(sys) 
sys.setdefaultencoding('utf-8')

def rmdir(top):
    for root, dirs, files in os.walk(top, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
    os.rmdir(top)

def zip_dir(dirname,zipfilename):
    filelist = []
    if os.path.isfile(dirname):
        filelist.append(dirname)
    else:
        for root, dirs, files in os.walk(dirname):
            for name in files:
                filelist.append(os.path.join(root, name))
    zf = zipfile.ZipFile(zipfilename, "w", zipfile.zlib.DEFLATED)
    for tar in filelist:
        arcname = tar[len(dirname):]
        zf.write(tar,arcname)
    zf.close()

def replace_file(filename, replace_dict):
    text = None
    with open(filename) as r:
        text = r.read()
        r.close()
    if text:
        for (keyword, value) in replace_dict.items():
            text = text.replace(keyword, value)
        with open(filename, 'w') as w:
            w.write(text)
            w.close()

def replace_with_template(template, docxname, project_no, project, project_en, date, date_en):
    f = zipfile.ZipFile(template)
    m2 = hashlib.md5()   
    m2.update(docxname)
    tmpdir = '.tmp' + m2.hexdigest()
    f.extractall(tmpdir)

    doc_dict = {
        '{project_no}': project_no,
        '{project}': project,
        '{project_en}': project_en,
        '{date}': date,
        '{date_en}': date_en
    }
    footer_dict = {
        '{project_no}': project_no
    }
    replace_file(tmpdir + '/word/document.xml', doc_dict)
    replace_file(tmpdir + '/word/footer1.xml', footer_dict)
    zip_dir(tmpdir, docxname)
    rmdir(tmpdir)

def process_excel_1(source, template, outdir):
    data = xlrd.open_workbook(source)
    table = data.sheet_by_index(0)
    for i in range(table.nrows):
        if i==0:
            continue
        col = table.row_values(i)
        replace_with_template(template, ''.join([outdir, '/','Section', str(col[0]), '.docx']), col[1], col[2], col[3], col[4], col[5])

#process_excel_1(u'sourcedata.xls', 'template.docx', '.')
