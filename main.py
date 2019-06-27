import xlrd
import xlwt
import math

class node:
    def __init__(self,x,y,z,r):
        self.x=x
        self.y=y
        self.z=z
        self.r=r
    def dist(self,other):
        return math.sqrt((self.x-other.x)**2+(self.y-other.y)**2+(self.z-other.z)**2)
    def neighbor(self,other):
        dist=self.dist(other)
        if self==other:
            return False
        if dist<self.r+other.r:
            return True
        else:
            return False

class cube:
    def __init__(self,xmin,xmax,ymin,ymax,zmin,zmax):
        self.xmin=xmin
        self.xmax=xmax
        self.ymin=ymin
        self.ymax=ymax
        self.zmin=zmin
        self.zmax=zmax

class square:
    def __init__(self,xmin,xmax,ymin,ymax):
        self.xmin=xmin
        self.xmax=xmax
        self.ymin=ymin
        self.ymax=ymax

class circle:
    def __init__(self,x,y,r):
        self.x=x
        self.y=y
        self.r=r

def read_excel():
    workbook=xlrd.open_workbook('聚集体形态.xlsx')
    sheet1=workbook.sheet_by_index(0)
    nodes=[]
    for i in range(sheet1.nrows):
        nodes.append(node(sheet1.cell(i,0).value,sheet1.cell(i,1).value,sheet1.cell(i,2).value,sheet1.cell(i,3).value))
    return nodes

def aggregate_border(nodes):
    xMin=float('inf')
    yMin=float('inf')
    zMin=float('inf')
    xMax=-float('inf')
    yMax=-float('inf')
    zMax=-float('inf')
    for n in nodes:
        if n.x-n.r<xMin:
            xMin=n.x-n.r
        if n.x+n.r>xMax:
            xMax=n.x+n.r
        if n.y-n.r<yMin:
            yMin=n.y-n.r
        if n.y+n.r>yMax:
            yMax=n.y+n.r
        if n.z-n.r<zMin:
            zMin=n.z-n.r
        if n.z+n.r>xMax:
            zMax=n.z+n.r
    return xMin, xMax, yMin, yMax, zMin, zMax

def build_cubes(xMin, xMax, yMin, yMax, zMin, zMax,len):
    cubes=[]
    xmax = xMin + len
    while xmax<xMax:
        ymax = yMin + len
        while ymax<yMax:
            zmax = zMin + len
            while zmax<zMax:
                cubes.append(cube(xmax-len,xmax,ymax-len,ymax,zmax-len,zmax))
                zmax=zmax+len
            ymax=ymax+len
        xmax=xmax+len
    return cubes

#判断正方形与圆形是否相交
def rectangleXcircle(rec,c):
    #第一种情况：圆心在长方形中间
    judge1 = c.x>=rec.xmin and c.x<=rec.xmax and c.y>=rec.ymin and c.y<=rec.ymax
    #第二种情况：圆心在长方形左下角
    judge2 = c.x < rec.xmin and c.y < rec.ymin and rec.xmin - c.x < c.r and rec.ymin-c.y < c.r
    # 第三种情况：圆心在长方形左上角
    judge3 = c.x < rec.xmin and c.y > rec.ymax and rec.xmin - c.x < c.r and c.y - rec.ymax < c.r
    # 第四种情况：圆心在长方形右下角
    judge4 = c.x > rec.xmax and c.y < rec.ymin and c.x - rec.xmax < c.r and rec.ymin - c.y < c.r
    # 第五种情况：圆心在长方形右上角
    judge5 = c.x > rec.xmax and c.y > rec.ymax and c.x - rec.xmax < c.r and c.y - rec.ymax < c.r
    #第六种情况：圆心在长方形在长方形左方
    judge6 = c.x < rec.xmin and c.y > rec.ymin and c.y <rec.ymax and rec.xmin - c.x <c.r
    # 第七种情况：圆心在长方形在长方形右方
    judge7 = c.x > rec.xmax and c.y > rec.ymin and c.y < rec.ymax and c.x - rec.xmax < c.r
    # 第八种情况：圆心在长方形在长方形上方
    judge8 = c.y > rec.ymax and c.x > rec.xmin and c.x < rec.xmax and c.y - rec.ymax < c.r
    # 第九种情况：圆心在长方形在长方形下方
    judge9 = c.y < rec.ymin and c.x > rec.xmin and c.x < rec.xmax and rec.ymin - c.y < c.r

    if judge1 or judge2 or judge3 or judge4 or judge5 or judge6 or judge7 or judge8 or judge9:
        return True
    else:
        return False

#判断球体与立方体是否相交
def cubeXnode(c,n):
    #投影到xoy平面，判断正方形与圆形是否相交
    judge1 = rectangleXcircle(square(c.xmin,c.xmax,c.ymin,c.ymax),circle(n.x,n.y,n.r))
    # 投影到xoz平面，判断正方形与圆形是否相交
    judge2 = rectangleXcircle(square(c.xmin, c.xmax, c.zmin, c.zmax), circle(n.x, n.z,n.r))
    # 投影到yoz平面，判断正方形与圆形是否相交
    judge3 = rectangleXcircle(square(c.ymin, c.ymax, c.zmin, c.zmax), circle(n.y, n.z,n.r))
    if judge1 and judge2 and judge3:
        return True
    else:
        return False

if __name__=='__main__':
    nodes=read_excel()
    xMin, xMax, yMin, yMax, zMin, zMax=aggregate_border(nodes)
    #print(xMin, xMax, yMin, yMax, zMin, zMax)
    min_range=min(xMax-xMin,yMax-yMin,zMax-zMin)
    #print(min_range)
    init=min_range/20.0
    step=min_range/20.0
    results={}
    length =init
    while length<min_range/2:
        cubes = build_cubes(xMin, xMax, yMin, yMax, zMin, zMax, length)
        #print(length,len(cubes))
        count = 0
        for c in cubes:
            #print(c.xmin,c.ymin,c.zmin)
            for n in nodes:
                if cubeXnode(c,n):
                    count = count + 1
                    break
        results[length]=count
        length=length+step
        #print(length,count)
    workbook=xlwt.Workbook(encoding='utf-8')
    worksheet=workbook.add_sheet('结果')
    worksheet.write(0,0,label='正方体长度')
    worksheet.write(0,1,label='正方体个数')
    row=1
    for len,count in results.items():
         worksheet.write(row,0,len)
         worksheet.write(row,1,count)
         row=row+1
    workbook.save('结果.xls')

