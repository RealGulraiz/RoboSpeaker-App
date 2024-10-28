import win32com.client as wincom

if __name__ == '__main__':
    speak = wincom.Dispatch("SAPI.SpVoice")
    text = input("Please Enter text to speak: ")
    speak.Speak(text)