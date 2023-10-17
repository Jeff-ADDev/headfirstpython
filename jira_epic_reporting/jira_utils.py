def test_zero_value(value, cell):
    if value == 0:
        cell.value = " - "
    else:
        cell.value = value
