import configparser
config = configparser.ConfigParser()

config.add_section('user_info')

config.set('user_info','filepath','/Users/jamesloh/PycharmProjects/WellesleyAutoSearch/datasets/data3.xlsx')
config.set('user_info','targetnamecolid','Name')
config.set('user_info','targetfirmcolid','Firm')
config.set('user_info','targettitlecolid','Title')
config.set('user_info','newfirmcol','E')
config.set('user_info','newtitlecol','F')
config.set('user_info','countsaved', '0')
#
# # Read config.ini file
# edit = configparser.ConfigParser()
# edit.read("config.ini")
# #Get the postgresql section
# user_info = edit["user_info"]
# #Update the password
# user_info["targetNameCol"] = "B"
# #Write changes back to file
with open('config.ini', 'w') as configfile:
    config.write(configfile)

# targetNameCol = 'B'
# targetFirmCol = 'A'
# newFirmCol = 'E'
# newTitleCol = 'F'
# countSaved = None
#
# print(targetNameCol)
