import unittest
from unittest.mock import patch, mock_open
from make_quantitative_bulk_load_file import *


class TestExtractValuesFromCSV(unittest.TestCase):
    csv_data = ['"indicator_id","indicator_name","country","year","quintile","service","value","real_value","currency","conversion_year","category","value_type","table_id","figure_id"',
                '"sel_annual","Mean annual subsistence expenditure line","SPA","2006","Total",NA,5789.83088716205,6955.25,"EUR","2020","Household budget survey","number","T1","F26"',
                '"poverty_line","Percent below subsistence expenditure line","SPA","2006","Total",NA,0.590924442657176,NA,"EUR","2020","Household budget survey","percentage","T1","F26"',
                '"ctp_annual","Mean annual capacity to pay","SPA","2006","Total",NA,24532.7152806292,29470.85,"EUR","2020","Household budget survey","number","T1","F26"',
                '"annual_oop_pc_quintile","Mean annual per capita OOP (by quintile)","SPA","2007","Poorest","NA","87.2864355545332","102.01","EUR","2020","Health spending","number","T2 Table 1","F5"',
                '"annual_oop_pc_quintile","Mean annual per capita OOP (by quintile)","SPA","2007","2nd","NA","190.509515508374","222.64","EUR","2020","Health spending","number","T2 Table 1","F5"',
                '"annual_oop_pc_quintile","Mean annual per capita OOP (by quintile)","SPA","2007","3rd","NA","272.954298750225","318.99","EUR","2020","Health spending","number","T2 Table 1","F5"',
                '"annual_oop_pc_quintile","Mean annual per capita OOP (by quintile)","SPA","2007","4th","NA","426.813408546795","498.81","EUR","2020","Health spending","number","T2 Table 1","F5"',
                '"annual_oop_pc_quintile","Mean annual per capita OOP (by quintile)","SPA","2007","Richest","NA","884.403235165156","1033.58","EUR","2020","Health spending","number","T2 Table 1","F5"']

    @patch('builtins.open', new_callable=mock_open, read_data='\n'.join(csv_data))
    def test_extract_values_from_csv(self, mock_open):

        from make_quantitative_bulk_load_file import extract_values_from_csv
        result = extract_values_from_csv('fake_file.csv')

        expected_result = {'Kingdom of Spain':
                           {
                               '2006': {
                                   'Mean monthly subsistence expenditure line (cost of meeting basic needs)': {
                                       'default': '5789.83088716205'
                                   },
                                   'Percent below subsistence expenditure line (basic needs line)': {
                                       'default': '0.590924442657176'
                                   },
                                   'Mean monthly capacity to pay for health care': {
                                       'default': '24532.7152806292'
                                   }
                               },
                               '2007': {
                                   "Annual out-of-pocket payments for health care per person (by consumption quintile)": {
                                       "Poorest": "87.2864355545332",
                                       "2nd": "190.509515508374",
                                       "3rd": "272.954298750225",
                                       "4th": "426.813408546795",
                                       "Richest": "884.403235165156"
                                   }
                               },
                           },
                           }

        self.assertEqual(result, expected_result)


class TestcreateDictIfDontExist(unittest.TestCase):
    def test_create_dict_if_dont_exist_existing_key(self):
        dictionary = {'key1': {'nested_key': 'nested_value'}}
        key = 'key1'
        create_dict_if_dont_exist(dictionary, key)
        self.assertEqual(dictionary, {'key1': {'nested_key': 'nested_value'}})

    def test_create_dict_if_dont_exist_new_key(self):
        dictionary = {'key1': {'nested_key': 'nested_value'}}
        key = 'key2'
        create_dict_if_dont_exist(dictionary, key)
        self.assertEqual(dictionary, {'key1': {'nested_key': 'nested_value'}, 'key2': {}})


class TestGetIndicator(unittest.TestCase):
    def test_get_indicator_id(self):
        ids = MetadataIds(
            countries={'Country1': 'Country1', 'Country2': 'Country2'},
            indicators={'Indicator1': 'Indicator1', 'Indicator2': 'Indicator2'},
            combos={'Combo1': 'Combo1', 'Combo2': 'Combo2'}
        )

        self.assertEqual(get_indicator_id(ids, 'Indicator1'), 'Indicator1')

        self.assertEqual(get_indicator_id(ids, 'dataElement'), None)

        self.assertEqual(get_indicator_id(ids, 'indicator3'), None)


class TestTestCheckMeanMonthlyIndicator(unittest.TestCase):
    def test_check_mean_monthly_indicator(self):
        self.assertEqual(check_mean_monthly_indicator(CTP_MONTHLY_NAME), True)

        self.assertEqual(check_mean_monthly_indicator("OTHER_NAME"), False)

        self.assertEqual(check_mean_monthly_indicator(""), False)


class TestMakeMatchedValues(unittest.TestCase):
    def test_make_matched_values(self):
        csv_values_dict = {
            'CTR1': {
                '2019': {
                    'IND1': {
                        'COC1': '10',
                        'COC2': '20'
                    },
                    'IND2': {
                        'COC1': '30',
                        'COC2': '40'
                    }
                },
                '2020': {
                    'IND1': {
                        'COC1': '50',
                        'COC2': '60'
                    },
                    'IND2': {
                        'COC1': '70',
                        'COC2': '80'
                    }
                }
            },
            'CTR2': {
                '2019': {
                    'IND1': {
                        'COC1': '90',
                        'COC2': '100'
                    },
                    'IND2': {
                        'COC1': '110',
                        'COC2': '120'
                    }
                },
                '2020': {
                    'IND1': {
                        'COC1': '130',
                        'COC2': '140'
                    },
                    'IND2': {
                        'COC1': '150',
                        'COC2': '160'
                    }
                }
            }
        }

        ids = MetadataIds(
            countries={'CTR1': 'Country1', 'CTR2': 'Country2'},
            indicators={'IND1': 'Indicator1', 'IND2': 'Indicator2'},
            combos={'Combo1': 'COC1', 'Combo2': 'COC2'}
        )

        expected_data = {
            'Country1': {
                '2019': {
                    'Indicator1': {
                        'Combo1': '10',
                        'Combo2': '20'
                    },
                    'Indicator2': {
                        'Combo1': '30',
                        'Combo2': '40'
                    }
                },
                '2020': {
                    'Indicator1': {
                        'Combo1': '50',
                        'Combo2': '60'
                    },
                    'Indicator2': {
                        'Combo1': '70',
                        'Combo2': '80'
                    }
                }
            },
            'Country2': {
                '2019': {
                    'Indicator1': {
                        'Combo1': '90',
                        'Combo2': '100'
                    },
                    'Indicator2': {
                        'Combo1': '110',
                        'Combo2': '120'
                    }
                },
                '2020': {
                    'Indicator1': {
                        'Combo1': '130',
                        'Combo2': '140'
                    },
                    'Indicator2': {
                        'Combo1': '150',
                        'Combo2': '160'
                    }
                }
            }
        }

        result = make_matched_values(csv_values_dict, ids)
        self.assertEqual(result, expected_data)


if __name__ == '__main__':
    unittest.main()
