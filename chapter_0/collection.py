import collections
how_many = "How many f's are in the following string? fffffffffffff"
c = collections.Counter(how_many)
print(c['f'])
