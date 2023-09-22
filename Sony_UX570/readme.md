
[Sony Device](https://www.amazon.com/dp/B082QL6KLG?psc=1&ref=ppx_yo2ov_dt_b_product_details)

# Installl These Packages

## Install openai
pip3 install openai

## Install python-dotenv
pip3 install python-dotenv

## Install PyDub
pip3 install pydub

## Install ffmpeg, find location, set location in script
brew install ffmpeg
/opt/homebrew/Cellar/ffmpeg/6.0_1/bin
from pydub import AudioSegment
AudioSegment.converter = '/opt/homebrew/Cellar/ffmpeg/6.0_1/bin'

## dotenv 
.env file in directory of file
### Contents
OPENAI_API_KEY=sk-Fe9U0ACXCSDxfiYkjPpyT3BlFgFXh0hox8dsTmsA3iWeYOPr
DIR_LOCATION=/Volumes/JEFFSPEECH/PRIVATE/SONY/REC_FILE/FOLDER01
COPY_AUDIO_TO_LOCATION=/users/jholmes/audionotes/audio
COPY_TEXT_TO_LOCATION=/users/jholmes/audionotes/text
FFMPEG_LOCATION=/opt/homebrew/Cellar/ffmpeg/6.0_1/bin/ffmpeg
TEMP_LOCATION=temp
SUMMARY_TIME=15
SEGMENT_TIME=600

And no that is not my OPENAI Key :-) 