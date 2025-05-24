import os
import json
import requests
import speech_recognition as s
import win32com.client
#start the app
if __name__=='__main__':
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    voices = speaker.GetVoices()
    speaker.Voice = voices.Item(1)
    speaker.speak('Welcome !!')
    speaker.Speak('i am listening you.......')
    #for always run i use while True
    while True:
        speaker.Speak('Tell your city name: ')
        sr=s.Recognizer()
        with s.Microphone() as m:
            audio=sr.listen(m)
            try:
                city=sr.recognize_google(audio,language='eng-in')
            except s.UnknownValueError:
                speaker.Speak("Sorry, I couldn't understand. Please try again.")
                continue
        # city=input("Enter city: ")
        if city.lower()=='exit':
            speaker.Speak('Thank you for using me.')
            break
        # speaker.Speak(query)
        url=f'https://api.weatherapi.com/v1/current.json?key=d71718881c614332827100118240804&q={city}'
        #requsting data from the api
        f=requests.get(url)
        dic=json.loads(f.text)
        district=dic['location']['region']
        country=dic['location']['country']
        speaker.speak(f'Are you asking about{city} {district} {country}')
        #try except for language understanding
        try:
            with s.Microphone() as m:
              confirm_audio=sr.listen(m)
              sr.adjust_for_ambient_noise(m)
              confirm=sr.recognize_google(confirm_audio,language='eng-in').lower()
        except s.UnknownValueError:
            speaker.Speak("Sorry, I couldn't understand. Please try again.")
            continue
        #city confirmetion from the user
        if confirm=='yes':
            speaker.Speak(f'what is you want to know in {city} related to weather?')
            #asking the what data user wants
            with s.Microphone() as m:

                try:
                    topic_audio=sr.listen(m)
                    topic=sr.recognize_google(topic_audio,language='eng-in').lower()
                except s.UnknownValueError:
                    speaker.Speak("Sorry, I couldn't understand your query. Please try again.")
                    continue

            #matching the condition given by user with api data
            if topic=='wind speed in meter per hours':
                windSpeedInMPH=f'Wind speed in {city} is {dic['current']['wind_mph']} meter per hours.'
                speaker.Speak(windSpeedInMPH)
            elif topic=='wind speed':
                windSpeed=f'Wind speed in {city} is {dic['current']['wind_kph']} kilometer per hours.'
                speaker.Speak(windSpeed)
            elif topic=='wind direction':
                wind_direction_map = {
                    'NE': 'North East',
                    'NS': 'North South',
                    'NW': 'North West',
                    'N': 'North',
                    'E': 'East',
                    'ES': 'East South',
                    'EW': 'East West',
                    'EN': 'East North',
                    'S': 'South',
                    'SW': 'South West',
                    'SN': 'South North',
                    'SE': 'South East',
                    'W': 'West',
                    'WN': 'West North',
                    'WE': 'West East',
                    'WS': 'West South'
                }
                dir=dic['current']['wind_dir']
                direction = wind_direction_map.get(dir, dir)
                windSpeed=f'Wind direction in {city} is {direction} .'
                speaker.Speak(windSpeed)
            elif topic=='pressure':
                press=f'pressure in {city} is {dic['current']['pressure_mb']} milibar.'
                speaker.Speak(press)
            elif topic=='real temperature':
                realTemp=f'Temperature feel in {city} is {dic['current']['feelslike_c']} degree centigrad.'
                speaker.Speak(realTemp)
            elif topic=='real temperature in fahrenhight':
                realTempInf=f'Temperature feel in {city} is {dic['current']['feelslike_c']} degree fahrenhight.'
                speaker.Speak(realTempInf)
            elif topic=='temperature':
                temp=f'weather in {city} is {dic['current']['temp_c']} degree centigrad.'
                speaker.Speak(temp)
            elif topic=='temperature in fahrenhight':
                temp=f'weather in {city} is {dic['current']['temp_f']} degree fahrenhight.'
                speaker.Speak(temp)
            elif topic=='humidity':
                humidity=f'Humidity in {city} is {dic['current']['humidity']} .'
                speaker.Speak(humidity)
            elif topic=='UV light':
                uv=f'Ualtra Visible light in {city} is {dic['current']['uv']} .'
                speaker.Speak(uv)
            else:
                speaker.Speak('I did not understand what you said. Please try again.')
        elif confirm=='no':
            speaker.Speak('Please tell your city name with the state and country.')

        #if we want to again run the 
        speaker.Speak('If you want to exit, say "exit". Otherwise')
