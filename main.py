import pandas as pd
import asyncio
from docx import Document
from googletrans import Translator



SITE_NAME = 'mikmebel'
FULL_SITE_NAME = 'mik-mebel.com'

# PATH_FILE = f'K:\\Work\\Audit_projects\\{FULL_SITE_NAME}\\tech-seo-audit-{SITE_NAME}.xlsx'
PATH_FILE = f'K:\\Work\\Audit_projects\\{FULL_SITE_NAME}\\tech-seo-audit-mikmebel.xlsx'
SHEET = '1 - Issues Overview'
DOC = 'K:\\Work\\Audit_projects\\ScreaF2Word\\template_base.docx'
DOC_NEW = f'K:\\Work\\Audit_projects\\ScreaF2Word\\tech-seo-audit-{SITE_NAME}.docx'
SRC = 'en'
DEST = 'ru'
TRAN = True

#def text_translator(text):
#    try:
#        translator = Translator()
#        trans_result = translator.translate(text, src=SRC, dest=DEST)
#        return trans_result.text
#
#    except Exception as ex:
#        return ex


async def text_translator(text, translator=None):
    try:
        translator = Translator(http2=True, timeout=5)
        trans_result = await translator.translate(text, src=SRC, dest=DEST)

        return trans_result.text
    except Exception as ex:
        return ex



def main():
    # file = pd.ExcelFile('K\\Work\\Audit_projects\\goodsnus.in.ua\\tech-seo-audit-goodsnus.xlsx')

    with pd.ExcelFile(PATH_FILE) as xlsx:
        df1 = pd.read_excel(xlsx, SHEET)

    print(df1.head(10))
    print(df1.size)
    print(df1.Description)

    doc = Document()

    idx = 1

    translator = Translator(http2=True, timeout=None)

    for issue, entry, priority, n_urls, h2_fix in zip(df1['Issue Name'], df1.Description, df1['Issue Priority'], df1['URLs'], df1['How To Fix']):
        print(f'Issue name: {issue}')
        print(f'(Issue Priority: {priority})')
        print(entry)

        txt_issue_name = f'{idx}. Problem Title: "{issue}"'
        txt_issue_priority = f'Problem priority: {priority}'
        txt_detected_urls = f'Detected URLs: {n_urls}'
        txt_entry = entry
        txt_fix_head = 'How to fix'
        txt_h2_fix = h2_fix

        # print(text_translator(txt_issue_name))

        if TRAN:
            txt_issue_name = asyncio.run(text_translator(txt_issue_name))
            txt_issue_priority = asyncio.run(text_translator(txt_issue_priority))
            txt_detected_urls = asyncio.run(text_translator(txt_detected_urls))
            txt_entry = asyncio.run(text_translator(txt_entry))
            txt_fix_head = asyncio.run(text_translator(txt_fix_head))
            txt_h2_fix = asyncio.run(text_translator(txt_h2_fix))

        print(txt_issue_name)
        print(txt_h2_fix)

        doc.add_heading(txt_issue_name, level=1)
        doc.add_heading(f'{txt_issue_priority} | {txt_detected_urls}', level=3)
        paragraph = doc.add_paragraph(txt_entry)
        paragraph.style.font.name = 'Calibri'
        doc.add_heading(txt_fix_head, level=3)
        doc.add_paragraph(txt_h2_fix)

        # doc.add_heading(f'{idx}. Issue name: {issue}', level=1)
        # doc.add_heading(f'Issue priority: {priority} | Detected URLs: {n_urls}', level=3)
        # paragraph = doc.add_paragraph(f'{entry}')
        # paragraph.style.font.name = 'Calibri'
        # doc.add_heading('How to fix', level=3)
        # doc.add_paragraph(f'{h2_fix}')

        idx += 1

    doc.save(DOC_NEW)


if __name__ == '__main__':
    main()
