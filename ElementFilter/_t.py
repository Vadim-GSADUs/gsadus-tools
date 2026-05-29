import importlib.util
spec=importlib.util.spec_from_file_location("ef","element_filter.py")
m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)

# parsing / joining
assert m.parse_ids("Elements of the current selection IDs are: 8744792, 8744643, 8745377")==[8744792,8744643,8745377]
assert m.parse_ids("100 100 200, 300; 200")==[100,200,300]
assert m.join_ids([1,2,3],"comma")=="1, 2, 3"
assert m.join_ids([1,2],"newline")=="1\n2"

# deduction: single poison narrowed by two overlapping locked tests
coll=[1,2,3,4,5,6]
log=[("dirty",[1,2,3,4]),("dirty",[3,4,5,6]),("clean",[3])]
d=m.deduce(coll,log)
assert d["clean_count"]==1
assert d["suspects"]==[1,2,4,5,6]            # 3 cleared
assert d["narrowed"]==[4]                    # only 4 is in both locked tests and not clean
assert d["confirmed"]==[]                    # not yet reduced to a size-1 locked test
assert d["multi"] is False

# a locked test reduced to one element => confirmed poison
d2=m.deduce([1,2,3],[("clean",[1,2]),("dirty",[2,3])])
assert d2["confirmed"]==[3]

# contradiction: locked test fully explained as clean => multiple poisons flag
d3=m.deduce([1,2,3],[("clean",[1,2,3]),("dirty",[1,2])])
assert d3["contradictions"]==[2]
assert d3["multi"] is True

# disjoint locked tests => no single culprit
d4=m.deduce([1,2,3,4],[("dirty",[1,2]),("dirty",[3,4])])
assert d4["narrowed"]==[]
assert d4["multi"] is True

print("ALL TESTS PASS")
