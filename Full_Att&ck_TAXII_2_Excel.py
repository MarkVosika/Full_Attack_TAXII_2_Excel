from stix2 import TAXIICollectionSource, Filter
from taxii2client.v20 import Server, Collection
import json
from openpyxl import Workbook
from openpyxl.styles import Font


#enterprise attack source
collection = Collection("https://cti-taxii.mitre.org/stix/collections/95ecc380-afe9-11e4-9b6c-751b66dd541e/")
# supply the TAXII2 collection to TAXIICollection
tc_source = TAXIICollectionSource(collection)

count = 1

def parse_mitre_item(filter_type):
	global count
	wb = Workbook()
	for item in filter_type:
		name = item[2]
		print (name)
		category = tc_source.query(item)


		#generate list of deduped list of keys
		key_list = []
		deduped_keys = []

		for lst in category:
			for k,v in lst.items():
				key_list.append(k)


		#remove duplicates and assign to list
		deduped_keys = sorted(list(dict.fromkeys(key_list)))


		#create a couple empty lists
		list_of_lists = []
		combined = []

		#create an empty list of lists, one for each dedupped item in deduped_keys
		for i in deduped_keys:
			list_name = []
			list_of_lists.append(list_name)

		#Loop though dedupped_key index and depending on the data type, parse through it if data exists, otherwise append "not found"	
		for technique in category:
				for index in enumerate(deduped_keys):
					if index[1] in technique:
						if isinstance(technique[index[1]], str):
							list_of_lists[index[0]].append(technique[index[1]])	
						elif isinstance(technique[index[1]], list):
							new_list = []
							for i in technique[index[1]]:
								if isinstance(i, str):
									x = i
									new_list.append(x)
								else:
									for k,v in i.items():
										x = k + ":" + v
										new_list.append(x)
							list_of_lists[index[0]].append(' | '.join(new_list))
						else:
							list_of_lists[index[0]].append(str(technique[index[1]]))
					else:
						list_of_lists[index[0]].append("Not Found")
		combined = zip(*list_of_lists)

		
		# write to excel file all the data, making first row bold
		bold_font = Font(bold = True)		#set bold font variable

		if count == 1:
			current_sheet = wb['Sheet']
			current_sheet.title = name
		else:
			current_sheet = wb.create_sheet(name)

		current_sheet.append(deduped_keys)
		for cell in current_sheet["1:1"]:
			cell.font = bold_font
		for i in combined:
			current_sheet.append(i)
		count += 1
	wb.save('all_mitre_parsed.xlsx')

#call the fuction with input types filter list
parse_mitre_item([Filter('type', '=', 'attack-pattern'), Filter('type', '=', 'intrusion-set'), Filter('type', '=', 'tool'), Filter('type', '=', 'malware'), Filter('type', '=', 'relationship')])
