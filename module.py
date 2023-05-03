glbVal = 0

class testCls:
  def setGlb(n):
    global glbVal
    glbVal = n

  def getGlb():
    global glbVal
    print(glbVal)