from configparser import ConfigParser
config_object = ConfigParser()

config_object['RTF TAGS'] = {
    "page break": "\\endnhere",
    "header": "\\header",
    "title": "\\trhdr",
    "row start": "\\trowd",
    "row end": "\\row",
    "cell end": "\\cell}"
}
config_object['RE EXPRESSIONS'] = {
    "font table": r"{\\fonttbl(.*)}",
    "font pattern": r'{\\f(\d+)\\.*? ([^;]+?);}',
    "header": r"{(?!\\)(.+)\\cell}",
    "headerstyle": r"(?<=q)[lrc]\\"
    
}
config_object['HEADER ALIGNMENT'] = {
    "l": "left",
    "c": "centre",
    "r": "right"
}

with open('config.ini', 'w') as conf:
    config_object.write(conf)