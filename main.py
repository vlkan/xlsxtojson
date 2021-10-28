import pandas
import json
import re
excel_data_df = pandas.read_excel('FILE_NAME', sheet_name='PAGE_NAME')  # utf8 OK
# Convert excel to string (define orientation of document in this case from up to down)
thisisjson = excel_data_df.to_json(orient='records')  # utf8 ERROR!
# Make the string into a list to be able to input in to a JSON-file
thisisjson_dict = json.loads(thisisjson)  # utf8 OK
length = len(thisisjson_dict)

# multi context fix
for i in range(0, length-1):
    temp = ""
    if (thisisjson_dict[i]['context'] != None):
        temp = thisisjson_dict[i]['context']
        # check next object. if it is None or Null, copy the previous one
        if (thisisjson_dict[i+1]['context'] == None):
            thisisjson_dict[i+1]['context'] = temp

# delete unnecessary char with regex
for i in range(0, length):
    text = thisisjson_dict[i]['context']
    pattern = r'\n'
    clean_text = re.sub(pattern, '', text)
    thisisjson_dict[i]['context'] = clean_text
# print(thisisjson_dict)

with open('Data.json', 'w', encoding='utf-8') as json_file:
    json.dump(thisisjson_dict, json_file, ensure_ascii=False)
