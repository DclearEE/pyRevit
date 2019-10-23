from itertools import imap

def tolist(obj1):
	if hasattr(obj1,"__iter__"): return obj1
	else: return [obj1]

elements = tolist(IN[0])
filter = map(str.lower, imap(str, tolist(IN[1]) ) )
in1, out1, nulls = [], [], []
OUT = in1, out1, nulls

for e in elements:
	c1 = UnwrapElement(e).Category
	if c1 is None:
		nulls.append(e)
	else:
		n1 = c1.Name.lower()
		if any(f in n1 for f in filter):
			in1.append(e)
		else:
			out1.append(e)