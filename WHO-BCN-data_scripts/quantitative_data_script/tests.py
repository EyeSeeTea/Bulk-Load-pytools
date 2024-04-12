import unittest
from unittest.mock import patch, mock_open
from make_quantitative_bulk_load_file import *




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
