import pandas as pd 

old = pd.read_csv('old_input_1217.csv')
new = pd.read_csv('new_input_1217.csv')
new.rename(columns={'*SSMax':'SSMax'}, inplace=True)
new.rename(columns={ new.columns[0]: "Date" },inplace=True)
old_columns = old.columns
new_columns = new.columns



def test_column_length():
    assert (len(old_columns)-len(new_columns)) == 0

def test_2():
    assert list(set(old_columns) - set(new_columns)) == "James Hare"





#We know that the new input has extra rows that go from end of day through midnight, this is an expected difference. 
