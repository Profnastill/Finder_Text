import shapely
import shapely as sh


pt1 = sh.Point([(9, 10)])
area2 = sh.Polygon([(9.0, 10.0), (9.0, 9.0), (9.5, 9.5)])

rt1=area2.intersection(pt1)


print(rt1)
if pt1== rt1:
    print(True)
else:
    print(False)


