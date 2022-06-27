import configparser
# Read config.ini file
path = '/Users/jamesloh/PycharmProjects/WellesleyAutoSearch/Demo/config.ini'
config = configparser.ConfigParser()
config.read(path)
#Get the postgresql section
user_info = config['user_info']
# #Update the password
# config.set('user_info','targetnamecolid','Coe')
# config.set('user_info','targetfirmcolid','cope')
# config.set('user_info','targettitlecolid','pog')
user_info["newfirmcol"] = "cope"
#Write changes back to file
with open(path, 'w') as configfile:
    config.write(configfile)

    "color:black;font-weight: 600;"
    "border: 2px solid black;"
    "border-radius: 5px;"
    "background-color: #EEF3F8;"
    "border-color: #007AFF"