#! /bin/bash

cd /home/user/reminder
echo "Starting Minecraft server..."
screen -dmS e-reminder /bin/bash -c "/usr/bin/python3 email-reminder.py" &
