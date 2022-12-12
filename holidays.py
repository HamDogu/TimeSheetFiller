from govuk_bank_holidays.bank_holidays import BankHolidays

# choose a different locale for holiday titles and notes
bank_holidays = BankHolidays(locale='en')
for bank_holiday in bank_holidays.get_holidays('england-and-wales', 2022):
    print(bank_holiday['title'], 'is on', bank_holiday['date'])
print(bank_holidays.get_next_holiday())
print()
