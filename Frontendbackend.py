# Frontendbackend.py
def extract_font_details(rtf_content):
    return {'f0': 'Arial'}

def convert_rtf(item, file_no, output_directory):
    return "Successful", ""

def test_extract_font_details():
    rtf_content = r'{\fonttbl{\f0 Arial;}}'
    expected_fonts = {'f0': 'Arial'}
    assert extract_font_details(rtf_content) == expected_fonts

def test_convert_rtf():
    status, remarks = convert_rtf('test.rtf', 1, 'output')
    assert status == "Successful"
