import sys, os
sys.stdout.reconfigure(encoding='utf-8')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'negostuy.settings')
import django; django.setup()

from docx import Document
from core.parser import extract_images_from_docx, collect_ordered_blocks
from core.kursovaya_classifier import classify_kursovaya_blocks
import tempfile, shutil

tmp = tempfile.mkdtemp()
try:
    imgs = extract_images_from_docx(r'C:/NeGostuy/неправильно.docx', tmp)
    doc = Document(r'C:/NeGostuy/неправильно.docx')
    blocks, paragraphs = collect_ordered_blocks(doc, imgs)
    elements = classify_kursovaya_blocks(blocks, paragraphs)
    print(f"Всего элементов: {len(elements)}\n")
    for e in elements:
        t = e.get('type', '')
        if t != 'paragraph':
            txt = (e.get('text') or '')[:70]
            print(f"[{t:25s}] {txt}")
finally:
    shutil.rmtree(tmp, ignore_errors=True)
