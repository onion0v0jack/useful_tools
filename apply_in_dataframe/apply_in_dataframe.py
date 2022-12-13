import pandas as pd

test_dict = {
    'x1': [1.230, 2.234, 3.34],
    'x2': [4.23462346, 5.99999994, 6.00000002],
    'x3': ['a', 'b', 'c'],
}
test_df = pd.DataFrame(test_dict)

###################################################################################

# apply in dataframe with one variable
def test_func_1(x, r = 3):
    # x: int float
    output = str(round(x, r))
    op_list = output.split('.')
    if len(op_list) > 1: # 小數點後面有東西
        return op_list[0] + '.' + op_list[-1] + '0' * (r - len(op_list[-1]))
    else:
        return output + '.' + '0' * r

# test_df['x1'].apply(test_func_1)   # 確定func中的變量只有1個時就直接apply
test_df.apply(lambda x: test_func_1(x['x1'], r = 4), axis = 1) # 若變量大於1個時就用multiple的方式，注意apply中lambda x的對象是誰

###################################################################################

# apply in dataframe with multiple variables
def test_func_2(x, y, r = 3):
    # x: int float     y:str
    output = str(round(x, r))
    op_list = output.split('.')
    if len(op_list) > 1: # 小數點後面有東西
        return y + op_list[0] + '.' + op_list[-1] + '0' * (r - len(op_list[-1]))
    else:
        return y + output + '.' + '0' * r

test_df.apply(lambda x: test_func_2(x['x1'], x['x3'], r = 4), axis = 1) # 這種方式是可以用於dateframe中多個column的
