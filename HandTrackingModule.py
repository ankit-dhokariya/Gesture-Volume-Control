import cv2
import mediapipe as mp
import time


class handDetector():
    def __init__(self, mode=False, maxHands=2, detectionCon=0.5, trackingCon=0.5):

        self.mode = mode
        self.maxHands = maxHands
        self.detectionCon = detectionCon
        self.trackingCon = trackingCon

        self.mphands = mp.solutions.hands
        self.hands = self.mphands.Hands(
            self.mode, self.maxHands, self.detectionCon, self.trackingCon)
        self.mpdraw = mp.solutions.drawing_utils

    def findHands(self, img, draw=True):

        rgb_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

        self.res = self.hands.process(rgb_img)

        if self.res.multi_hand_landmarks:
            for handLandmark in self.res.multi_hand_landmarks:
                if draw:
                    self.mpdraw.draw_landmarks(
                        img, handLandmark, self.mphands.HAND_CONNECTIONS)

        return img

    def findPosition(self, img, handNo=0, draw=True):

        landmarksList = []

        if self.res.multi_hand_landmarks:
            myHand = self.res.multi_hand_landmarks[handNo]

            for i, landmark in enumerate(myHand.landmark):

                h, w, c = img.shape
                cx, cy = int(landmark.x * w), int(landmark.y * h)
                landmarksList.append([i, cx, cy])
                if draw:
                    cv2.circle(img, (cx, cy), 5, (255, 0, 0), cv2.FILLED)

        return landmarksList


def main():
    currTime = 0
    prevTime = 0

    cap = cv2.VideoCapture(0)

    detector = handDetector()

    while True:
        _, img = cap.read()

        img = detector.findHands(img)
        landmarksList = detector.findPosition(img, draw=True)

        if len(landmarksList) != 0:
            print(landmarksList[4])

        currTime = time.time()
        fps = 1 / (currTime - prevTime)
        prevTime = currTime

        cv2.putText(img, f"FPS: {int(fps)}", (20, 70),
                    cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 0), 3)

        cv2.imshow("Hand Tracking", img)
        cv2.waitKey(1)


if __name__ == "__main__":
    main()
