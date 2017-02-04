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

def replace_with_template_dict(template, docxname, replace_dict):
    f = zipfile.ZipFile(template)
    m2 = hashlib.md5()   
    m2.update(docxname)
    tmpdir = '.tmp' + m2.hexdigest()
    f.extractall(tmpdir)
    replace_file(tmpdir + '/word/document.xml', replace_dict)   # 替换word文档主要文本内容
    replace_file(tmpdir + '/word/footer1.xml', replace_dict)    # 替换脚注
    zip_dir(tmpdir, docxname)
    rmdir(tmpdir)

def process_excel_1(source, template, outdir):
    data = xlrd.open_workbook(source)
    table = data.sheet_by_index(0)
    for i in range(table.nrows):
        if i==0:
            continue
        col = table.row_values(i)
        if str(col[0]).isdigit(): ##  序号为数字时才处理
            replace_with_template(template, ''.join([outdir, '/','Section', str(col[0]), '.docx']), col[1], col[2], col[3], col[4], col[5])

#process_excel_1(u'sourcedata.xls', 'template.docx', '.')

def process_with_config(config_file_path, outdir):
    config = None
    # 读取配置文档
    with open(config_file_path, 'r') as config_file:
        config = json.loads(str(config_file.read()), encoding='utf-8')
        config_file.close()
    if not config:
        return
    for task in config:
        source = task['source']
        templates = task['templates']
        keywords = task['keywords']

        data = xlrd.open_workbook(source)   # 打开Excel源数据
        table = data.sheet_by_index(0)

        dict_with_id = {}
        for keyword in keywords:            # 生成关键词与Excel表格列号对应关系
            dict_with_id[ keyword['string'] ] = keyword['index']

        for i in range(table.nrows):
            if i==0:
                continue
            col = table.row_values(i)
            if str(col[0]).isdigit():       #  序号为数字时才处理
                dict_replace = {}           #  生成关键词替换字典
                for (keystring, col_id) in dict_with_id.items():
                    dict_replace[keystring] = str(col[col_id])
                # 执行模版doc关键词替换
                for template in templates:
                    replace_with_template_dict(
                        template['file'], 
                        ''.join([outdir, '/', template['name'],'.docx']), 
                        dict_replace)





#process_with_config('template.json', 'demo')