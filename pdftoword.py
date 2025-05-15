#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Created by xiaoqin00 on 2017/6/26

#pdf 转为word,没有找到pdf直接转换为word的方法，就先转为txt，然后转换为word

import sys
import os
from pdfminer.pdfinterp import PDFResourceManager,PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

from optparse import OptionParser
from docx import Document
from docx.shared import Inches

#main
def pdftotxt(pdf_file):
    #输出文件名
    outfile = pdf_file + '.txt'

    debug = 0
    pagenos = set()
    password = ''
    maxpages = 0
    rotation = 0
    codec = 'utf-8'   #输出编码
    caching = True
    imagewriter = None
    laparams = LAParams()
    
    try:
        PDFResourceManager.debug = debug
        PDFPageInterpreter.debug = debug

        rsrcmgr = PDFResourceManager(caching=caching)
        outfp = open(outfile, 'w', encoding=codec)
        #pdf转换
        device = TextConverter(rsrcmgr, outfp, codec=codec, laparams=laparams,
                    imagewriter=imagewriter)

        fp = open(pdf_file, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        #处理文档对象中每一页的内容
        for page in PDFPage.get_pages(fp, pagenos,
                          maxpages=maxpages, password=password,
                          caching=caching, check_extractable=True) :
            page.rotate = (page.rotate+rotation) % 360
            interpreter.process_page(page)
        fp.close()
        device.close()
        outfp.close()
        
        # 检查生成的txt文件是否为空
        if os.path.getsize(outfile) == 0:
            print(f"警告: 生成的文本文件为空，PDF可能受保护或无法提取文本")
            return False
        return True
    except Exception as e:
        print(f"转换PDF到TXT时出错: {e}")
        return False

def txttoword(txt_file, output_file=None):
    #创建 Document 对象，相当于打开一个 word 文档
    document = Document()

    try:
        with open(txt_file, 'r', encoding='utf-8') as f:
            content = f.read()
            if not content.strip():
                print(f"警告: 文本文件 {txt_file} 为空，无法转换为Word")
                return False
                
            for i in f.readlines():
                i = i.strip()
                if not i:
                    i = '\t'
                p = document.add_paragraph(i)
        
        #保存文本
        if output_file:
            document.save(output_file)
        else:
            document.save(txt_file.replace('.txt', '.docx'))
        return True
    except Exception as e:
        print(f"转换TXT到Word时出错: {e}")
        return False

# 添加直接从PDF转换为Word的函数
def pdftoword(pdf_file, output_file=None):
    try:
        from pdf2docx import Converter
        
        if output_file is None:
            output_file = pdf_file.replace('.pdf', '.docx')
            
        # 使用pdf2docx库直接转换
        cv = Converter(pdf_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()
        
        print(f"已直接将PDF转换为Word: {output_file}")
        return True
    except ImportError:
        print("未安装pdf2docx库，尝试使用替代方法...")
        # 如果没有pdf2docx库，使用原来的两步转换方法
        txt_file = pdf_file + '.txt'
        if pdftotxt(pdf_file):
            result = txttoword(txt_file, output_file)
            # 删除中间的TXT文件
            if os.path.exists(txt_file):
                os.remove(txt_file)
            return result
        return False
    except Exception as e:
        print(f"直接转换PDF到Word时出错: {e}")
        return False

def process_all_pdfs(directory=None):
    # 获取指定目录或当前目录下所有的PDF文件
    if directory is None:
        # 如果没有指定目录，使用脚本所在的目录
        directory = os.path.dirname(os.path.abspath(__file__))
    
    print(f"在目录 {directory} 中查找PDF文件...")
    
    try:
        pdf_files = [f for f in os.listdir(directory) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            print(f"在 {directory} 目录下没有找到PDF文件")
            
            # 尝试在上级目录中查找
            parent_dir = os.path.dirname(directory)
            print(f"尝试在上级目录 {parent_dir} 中查找...")
            pdf_files = [f for f in os.listdir(parent_dir) if f.lower().endswith('.pdf')]
            
            if pdf_files:
                directory = parent_dir
                print(f"在上级目录中找到 {len(pdf_files)} 个PDF文件")
            else:
                print(f"在上级目录中也没有找到PDF文件")
                return
        
        print(f"找到 {len(pdf_files)} 个PDF文件:")
        for pdf in pdf_files:
            print(" - " + pdf)
        
        print("\n开始转换...")
        success_count = 0
        
        for pdf_file in pdf_files:
            pdf_path = os.path.join(directory, pdf_file)
            print(f"处理: {pdf_file}")
            
            # 使用直接转换方法
            output_path = pdf_path.replace('.pdf', '.docx')
            if pdftoword(pdf_path, output_path):
                success_count += 1
                print(f"完成: {pdf_file} -> {os.path.basename(output_path)}")
        
        print(f"\n转换完成! 成功转换 {success_count}/{len(pdf_files)} 个文件")
    
    except Exception as e:
        print(f"查找或处理PDF文件时出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"正在处理脚本所在目录: {script_dir}")
    
    try:
        # 获取脚本目录下所有的PDF文件
        pdf_files = [f for f in os.listdir(script_dir) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            print("脚本所在目录下没有找到PDF文件")
            sys.exit(1)
            
        print(f"找到 {len(pdf_files)} 个PDF文件:")
        for pdf in pdf_files:
            print(" - " + pdf)
        
        print("\n开始转换...")
        success_count = 0
        
        for pdf_file in pdf_files:
            pdf_path = os.path.join(script_dir, pdf_file)
            print(f"处理: {pdf_file}")
            
            # 使用直接转换方法
            output_path = pdf_path.replace('.pdf', '.docx')
            if pdftoword(pdf_path, output_path):
                success_count += 1
                print(f"完成: {pdf_file} -> {os.path.basename(output_path)}")
        
        print(f"\n转换完成! 成功转换 {success_count}/{len(pdf_files)} 个文件")
    
    except Exception as e:
        print(f"处理PDF文件时出错: {e}")
        import traceback
        traceback.print_exc()