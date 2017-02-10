# encoding: utf-8
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

def replace_with_template_dict(template, outdir, docxname, replace_dict):
    f = zipfile.ZipFile(template)
    m2 = hashlib.md5()
    full_out_doc_path = outdir + docxname
    m2.update(full_out_doc_path)
    tmpdir = '.tmp' + m2.hexdigest()
    if os.path.exists(tmpdir):
        rmdir(tmpdir)
    f.extractall(tmpdir)
    if os.path.exists(tmpdir + '/word/document.xml'):
        replace_file(tmpdir + '/word/document.xml', replace_dict)   # 替换word文档主要文本内容
    if os.path.exists(tmpdir + '/word/footer1.xml'):
        replace_file(tmpdir + '/word/footer1.xml', replace_dict)    # 替换脚注
    if not os.path.exists(outdir):
        os.makedirs(outdir)
    zip_dir(tmpdir, os.path.join(outdir, docxname))
    rmdir(tmpdir)

def isNum(value):
    try:
        value * 1
    except TypeError:
        return False
    else:
        return True

def process_with_config(config_file_path, outdir):
    config = None
    config_base_dir = os.path.dirname(config_file_path)
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

        data = xlrd.open_workbook(''.join([config_base_dir,'/',source]))   # 打开Excel源数据
        table = data.sheet_by_index(0)

        dict_with_id = {}
        unique_key_id = 0
        for keyword in keywords:            # 生成关键词与Excel表格列号对应关系
            if 'unique' in keyword:
                unique_key_id = keyword['index'] # 找到唯一的Key，用于项目文件夹命名（这里是项目编号）
            dict_with_id[ keyword['string'] ] = keyword['index']
        
        for i in range(table.nrows):
            if i==0:
                continue
            col = table.row_values(i)
            if not isNum(col[0]) or not col[0]:
                continue
            #  序号为数字时才处理
            dict_replace = {}           #  生成关键词替换字典
            for (keystring, col_id) in dict_with_id.items():
                dict_replace[keystring] = str(col[col_id])

            # 生成单条数据文件夹名字
            item_name = str(col[unique_key_id])
            if 'item_name_rule' in task: # 如果有命名规则，则使用命名规则替换
                item_name = task['item_name_rule'].replace('{unique_key}', str(col[unique_key_id]))
            item_path = os.path.join(outdir, task['project'], item_name)
            # 执行模版doc关键词替换
            for template in templates:
                # 如果存在需要被替换的word文档则替换
                out_doc_name = None
                out_dir = os.path.join(item_path, template['outdir'])
                if not os.path.exists(out_dir):
                    os.makedirs(out_dir)
                if 'file' in template:
                    out_doc_name = template['name'] + '.docx'
                    in_doc_name = os.path.join(config_base_dir, template['file'])
                    replace_with_template_dict(
                        in_doc_name,    # 读入的模板word文档
                        out_dir,         # word输出的目录
                        out_doc_name,   # 输出的word名字
                        dict_replace)
                # 复制文件夹内容
                if 'indir' in template and template['indir']:
                    indir = os.path.join(config_base_dir, template['indir'])
                    for f in os.listdir(indir):  
                        if not cmp(f, out_doc_name): # 防止文件被覆盖
                            continue
                        sourceF = os.path.join(indir, f)
                        if os.path.isfile(sourceF):
                            if not os.path.exists(out_dir):
                                os.makedirs(out_dir)
                            targetF = os.path.join(out_dir, f)
                            open(targetF, "wb").write(open(sourceF, "rb").read())

            
            final_zip_name = item_name + '.zip'
            zip_dir(item_path, final_zip_name)                          # 压缩打包


process_with_config(u'C:/Users/tomic/Desktop/新建文件夹/template.json', '.')
