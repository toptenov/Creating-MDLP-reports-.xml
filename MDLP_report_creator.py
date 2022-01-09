import os
import openpyxl
import xml.etree.ElementTree as ET
import datetime

CURRENT_DATE = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S') 


def create_result_directory():
	if not os.path.exists("Result"):
		os.mkdir("Result")


def read_source_data():
	data = [{"sscc": []}]
	report_number = input("Введите номер формируемого МДЛП-отчёта (по умолчанию 415): ")

	# Делаем проверку на пустой ввод report_number
	if report_number == "":
		report_number = "415"

	book = openpyxl.load_workbook("Source.xlsx")
	sheet = book[report_number]

	# Собираем SSCC:
	for row in range(2, len(sheet["A"])+1):
		sscc = sheet[row][0].value
		data[0]["sscc"].append(sscc)

	# Собираем остальные данные:
	data[0]["subject_id"] = sheet[2][1].value
	data[0]["receiver_id"] = sheet[2][2].value
	data[0]["operation_date"] = sheet[2][3].value
	data[0]["doc_num"] = sheet[2][4].value
	data[0]["doc_date"] = sheet[2][5].value
	data[0]["turnover_type"] = sheet[2][6].value
	data[0]["source"] = sheet[2][7].value
	data[0]["contract_type"] = sheet[2][8].value
	data[0]["cost"] = sheet[2][9].value
	data[0]["vat_value"] = sheet[2][10].value

	return data, report_number


def create_xml(data, report_number):
	documents = ET.Element("documents")
	move_order = ET.SubElement(documents, "move_order", action_id=f"{report_number}")
	ET.SubElement(move_order, "subject_id").text = data[0]["subject_id"]
	ET.SubElement(move_order, "receiver_id").text = data[0]["receiver_id"]
	ET.SubElement(move_order, "operation_date").text = data[0]["operation_date"]
	ET.SubElement(move_order, "doc_num").text = data[0]["doc_num"]
	ET.SubElement(move_order, "doc_date").text = data[0]["doc_date"]
	ET.SubElement(move_order, "turnover_type").text = data[0]["turnover_type"]
	ET.SubElement(move_order, "source").text = data[0]["source"]
	ET.SubElement(move_order, "contract_type").text = data[0]["contract_type"]
	order_details = ET.SubElement(move_order, "order_details")

	for i in range(len(data[0]["sscc"])):
		union = ET.SubElement(order_details, "union")
		sscc_detail = ET.SubElement(union, "sscc_detail")
		ET.SubElement(sscc_detail, "sscc").text = data[0]["sscc"][i]
		ET.SubElement(union, "cost").text = data[0]["cost"]
		ET.SubElement(union, "vat_value").text = data[0]["vat_value"]

	tree = ET.ElementTree(documents)
	ET.indent(tree, space="    ")
	tree.write(f"./Result/{CURRENT_DATE}_{report_number}.xml", encoding="UTF-8", xml_declaration=True)


def main():
	create_result_directory()
	data, report_number = read_source_data()
	create_xml(data, report_number)


if __name__ == '__main__':
	main()
