import asyncio
import deepl
import os
import time
import logging
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

deepl_client = deepl.DeepLClient(auth_key)


class SourceLangError(Exception):
	def __str__(self):
		return 'The detected language does not match specified source!'


def _sync_translate_worker(text: str) -> str:
	"""
	Эта функция выполняет синхронный запрос.
	Она будет вызываться в отдельном потоке, чтобы не блокировать главный цикл.
	"""
	trans_result = deepl_client.translate_text(text, source_lang=SRC, target_lang=DEST)

	if trans_result.detected_source_lang != SRC:
		raise SourceLangError()

	return trans_result.text


async def text_translator(text):
	"""
	    Асинхронная функция, которую мы будем вызывать с await.
	    Она делегирует тяжелую сетевую задачу в фоновый поток.
	    """
	try:
		# asyncio.to_thread запускает синхронную функцию в отдельном потоке
		# и возвращает управление главному циклу, пока поток работает
		result = await asyncio.to_thread(_sync_translate_worker, text)
		return result

	except SourceLangError as err:
		print(f"⚠️ Ошибка языка: {err}. Текст оставлен без изменений.")
		return text  # Возвращаем исходный текст, чтобы не ломать документ
	except Exception as ex:
		print(f"❌ Ошибка перевода: {ex}. Текст оставлен без изменений.")
		return text  # Возвращаем исходный текст для надежности


async def main():
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

			# task1 = asyncio.create_task(text_translator(txt_issue_name))
			# task2 = asyncio.create_task(text_translator(txt_issue_priority))
			# task3 = asyncio.create_task(text_translator(txt_detected_urls))
			# task4 = asyncio.create_task(text_translator(txt_entry))
			# task5 = asyncio.create_task(text_translator(txt_fix_head))
			# task6 = asyncio.create_task(text_translator(txt_h2_fix))

			results = await asyncio.gather(
				text_translator(txt_issue_name),
				text_translator(txt_issue_priority),
				text_translator(txt_detected_urls),
				text_translator(txt_entry),
				text_translator(txt_fix_head),
				text_translator(txt_h2_fix)
			)

			txt_issue_name, txt_issue_priority, txt_detected_urls, txt_entry, txt_fix_head,	txt_h2_fix = results

			await asyncio.sleep(0.5)

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
	asyncio.run(main())
