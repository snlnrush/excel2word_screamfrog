import deepl
import time
import logging
import os
import pandas as pd
from docx import Document
from dotenv import load_dotenv


SITE_NAME = 'welwet'
FULL_SITE_NAME = 'welwet.net'

# PATH_FILE = f'K:\\Work\\Audit_projects\\{FULL_SITE_NAME}\\tech-seo-audit-{SITE_NAME}.xlsx'
PATH_FILE = f'E:\\Work\\Audit_projects\\{FULL_SITE_NAME}\\tech-seo-audit-welwet.xlsx'
SHEET = '1 - Issues Overview'
DOC = 'E:\\Work\\Audit_projects\\ScreaF2Word\\template_base.docx'
DOC_NEW = f'E:\\Work\\Audit_projects\\ScreaF2Word\\tech-seo-audit-{SITE_NAME}.docx'
SRC = 'EN'
DEST = 'RU'
TRAN = True

logging.basicConfig()
logging.getLogger('deepl').setLevel(logging.DEBUG)

load_dotenv()

auth_key = os.getenv("API_KEY")  # Replace with your key


class SourceLangError(Exception):
	def __str__(self):
		return 'The detected language does not match specified source!'



def text_translator(text):
	try:
		deepl_client = deepl.DeepLClient(auth_key)
		trans_result = deepl_client.translate_text(text, source_lang=SRC, target_lang=DEST)

		if trans_result.detected_source_lang != SRC:
			raise SourceLangError

		return trans_result.text
	except SourceLangError as err:
		print(err)
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

    for issue, entry, priority, n_urls, h2_fix in zip(df1['Issue Name'], df1.Description, df1['Issue Priority'], df1['URLs'], df1['How To Fix']):
        print(f'Issue name: {issue}')
        print(f'(Issue Priority: {priority})')
        print(entry)

        txt_issue_name = f'{idx}. Issue name: "{issue.capitalize()}"'
        txt_issue_priority = f'Issue priority: {priority.upper()}'
        txt_detected_urls = f'Detected URLs: {n_urls}'
        txt_entry = entry
        txt_fix_head = 'How to fix'
        txt_h2_fix = h2_fix

        # print(text_translator(txt_issue_name))

        if TRAN:

            task1 = text_translator(txt_issue_name)
            task2 = text_translator(txt_issue_priority)
            task3 = text_translator(txt_detected_urls)
            task4 = text_translator(txt_entry)
            task5 = text_translator(txt_fix_head)
            task6 = text_translator(txt_h2_fix)

            time.sleep(1)

            txt_issue_name = task1
            txt_issue_priority = task2
            txt_detected_urls = task3
            txt_entry = task4
            txt_fix_head = task5
            txt_h2_fix = task6

        print(txt_issue_name)
        print(txt_h2_fix)

        doc.add_heading(txt_issue_name, level=1)
        doc.add_heading(f'{txt_issue_priority} | {txt_detected_urls}', level=3)
        paragraph = doc.add_paragraph(txt_entry)
        paragraph.style.font.name = 'Calibri'
        doc.add_heading(txt_fix_head, level=3)
        doc.add_paragraph(txt_h2_fix)

        idx += 1

    doc.save(DOC_NEW)


if __name__ == '__main__':
    main()
