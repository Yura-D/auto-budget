
def get_next_year_month(year, month):
    """Return next_year, next_month"""
    if month == 12:
        next_month = 1
        next_year = year + 1
    else:
        next_month = month + 1
        next_year = year
    return next_year, next_month


def get_use_left(year, month, row_number):
    next_year, next_month = get_next_year_month(year, month)

    USE_LEFT = {
        "USE": f"=SUM(FILTER('Form Responses 1'!D:D, 'Form Responses 1'!B:B>\
            DATE({year},{month},5),'Form Responses 1'!B:B<\
                DATE({next_year},{next_month},6)))",
        
        "YURA_USE": f"=sum(FILTER('Form Responses 1'!D:D, 'Form Responses 1'\
            !E:E=\"Yura\",'Form Responses 1'!B:B>DATE({year},{month},5),\
                'Form Responses 1'!B:B<DATE({next_year},{next_month},6)))",
        
        "TANYA_USE": f"=sum(FILTER('Form Responses 1'!D:D, 'Form Responses 1'\
            !E:E=\"Tanya\",'Form Responses 1'!B:B>DATE({year},{month},5),\
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


def fill_template(category, year_start, year_end, month_start, month_end):
    category = category.capitalize()
    template = (
        "=sum(FILTER('Form Responses 1'!D:D,'Form Responses " +
        f"1'!C:C=\"{category}\",'Form Responses 1'!B:B>" +
        f"DATE({year_start},{month_start},5),'Form Responses 1'" +
        f"!B:B<DATE({year_end},{month_end},6)))"
    )
    return template


def get_statistics(year, month, row_number):
    next_year, next_month = get_next_year_month(year, month)
    CATEGORIES = [
        "FOODSTUFF",
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
        c: fill_template(c, year, next_year, month, next_month)
            for c in CATEGORIES
    }
    formulas['SUM'] = f'=SUMIF(B{row_number}:J{row_number},"<>#N/A")'
    formulas['VACATION'] = ''
    formulas['SUPER_SUM'] = F'=K{row_number}+L{row_number}'

    return formulas
