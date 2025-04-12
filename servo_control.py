import board
import busio
from adafruit_servokit import ServoKit
from time import sleep

i2c = busio.I2C(board.SCL, board.SDA, frequency=400000)
kit = ServoKit(channels=16, i2c=i2c)

# Move only servos 0, 1, 2, 3, and 4 continuously
def move_active_servos():
    while True:
        for angle in range(0, 180, 5):  # Move forward
            for i in [0, 1, 2, 3, 4]:  # Only move plugged-in servos
                kit.servo[i].angle = angle
            sleep(0.02)  # Adjust speed

        for angle in range(180, 0, -5):  # Move backward
            for i in [0, 1, 2, 3, 4]:
                kit.servo[i].angle = angle
            sleep(0.02)

# Run the loop continuously
move_active_servos()
