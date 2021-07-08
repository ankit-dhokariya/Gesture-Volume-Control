import cv2
import time
import numpy as np
import HandTrackingModule as htm
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

detector = htm.handDetector(detectionCon=0.8, maxHands=1)
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

################################ Paramters ################################
camWidth, camHeight = 640, 480
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]
volBarHeight = 300
volBarWidth = 35
vol = volume.GetMasterVolumeLevel()
volPer = np.interp(vol, [-63.5, 0], [0, 100])
volBar = np.interp(volPer, [0, 100], [100 + volBarHeight, 100])
###########################################################################

cap = cv2.VideoCapture(0)
cap.set(3, camWidth)
cap.set(4, camHeight)
prevTime = 0


while True:
    success, img = cap.read()
    if success:
        img = detector.findHands(img)
        landmarksList = detector.findPosition(img, draw=False)
        if len(landmarksList):

            x1, y1 = landmarksList[4][1], landmarksList[4][2]
            x2, y2 = landmarksList[8][1], landmarksList[8][2]
            cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

            cv2.circle(img, (x1, y1), 7, (255, 0, 255), cv2.FILLED)
            cv2.circle(img, (x2, y2), 7, (255, 0, 255), cv2.FILLED)
            cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 2)
            cv2.circle(img, (cx, cy), 7, (255, 0, 255), cv2.FILLED)

            length = math.hypot(x2 - x1, y2 - y1)
            vol = np.interp(length, [15, 130], [minVol, maxVol])
            volPer = np.interp(vol, [-63.5, 0], [0, 100])
            volBar = np.interp(volPer, [0, 100], [100 + volBarHeight, 100])

            volume.SetMasterVolumeLevel(int(vol), None)

            if length < 15:
                cv2.circle(img, (cx, cy), 7, (0, 255, 0), cv2.FILLED)

        cv2.rectangle(img, (50, 100), (50 + volBarWidth,
                                       100 + volBarHeight), (255, 0, 0), 3)
        cv2.rectangle(img, (50, int(volBar)), (50 + volBarWidth,
                                               100 + volBarHeight), (0, 0, 255), cv2.FILLED)
        cv2.putText(img, f'{int(volPer)}%', (20, 150 + volBarHeight),
                    cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 2)

        currTime = time.time()
        fps = 1 / (currTime - prevTime)
        prevTime = currTime
        cv2.putText(img, f'FPS: {int(fps)}', (20, 50),
                    cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 2)

        cv2.imshow("Image", img)
        cv2.waitKey(1)
