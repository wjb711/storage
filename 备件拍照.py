import cv2
import easygui as g
def enter_ID():
    
    id=g.enterbox(title="物品编码")
    return id
cap=cv2.VideoCapture(1)
while True:
    _,frame=cap.read()
    cv2.imshow("Press Enter for capture",frame)
    if cv2.waitKey(1)==13:
        id=enter_ID()
        cv2.imwrite(id+'.jpg',frame)
        
        
    if cv2.waitKey(1)==27:
        break
