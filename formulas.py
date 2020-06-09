
def get_use_left(year, month, row_number):
    if month == 12:
        next_month = 1
        next_year = year + 1
    else:
        next_month = month + 1
        next_year = year

    USE_LEFT = {
        "USE": f"=SUM(FILTER('Form Responses 1'!D:D, 'Form Responses 1'!B:B>\
            DATE({year},{month},5),'Form Responses 1'!B:B<\
                DATE({next_year},{next_month},6)))",
        
        "YURA_USE": f"=sum(FILTER('Form Responses 1'!D:D, 'Form Responses 1'\
            !E:E=\"Yura\",'Form Responses 1'!B:B>DATE({year},{month},5),\
                'Form Responses 1'!B:B<DATE({next_year},{next_month},6)))",
        
        "TANYA_USE": f"=sum(FILTER('Form Responses 1'!D:D, 'Form Responses 1\
            '!E:E=\"Tanya\",'Form Responses 1'!B:B>DATE({year},{month},5),\
                'Form Responses 1'!B:B<DATE({next_year},{next_month},6)))",
        
        "YURA_LEFT": f"=G{row_number}*0.6-C{row_number}",
        "TANYA_LEFT": f"=G{row_number}*0.4-D{row_number}",
        "FOND": "8000",
        "PLAN": "8000",
        "COMMENT": "",
        "EMPTY": "",
        "YURA_NEED": f"=B{row_number}*0.6",
        "TANYA_NEED": f"=B{row_number}*0.4"
    }
    return USE_LEFT



def get_statistics(year_start, year_end, month_start, month_end):
    TEMPLATE = (
        "=sum(FILTER('Form Responses 1'!D:D,'Form Responses " +
        "1'!C:C=\"{CATEGORY}\",'Form Responses 1'!B:B>" +
        "DATE({YEAR_START},{MONTH_START},5),'Form Responses 1'" +
        "!B:B<DATE({YEAR_END},{MONTH_END},6)))"
    )
    CATEGORIES = [
        "FOODSTAFF",
        "REST",
        "SWEET",
        "CAFE",
        "HYGIENE",
        "HOME",
        "BILLS",
        "CAT",
        "OTHER"
    ]
    formulas = {
        c: TEMPLATE.replace('{CATEGORY}', c.capitalize())\
            .replace('{YEAR_START}', str(year_start))\
                .replace('{MONTH_START}', str(month_start))\
                    .replace('{YEAR_END}', str(year_end))\
                        .replace('{MONTH_END}', str(month_end))\
                            for c in CATEGORIES
    }
    return formulas