import speech_recognition as sr
from pydub import AudioSegment


def quebra_captcha():
    song = AudioSegment.from_wav(r"E:\codigos_git\Novas-Tecnologias-Barbieri\Correcao EPROC\som_captcha.wav")
    beginning = song + 13

    beginning.export("teste.wav", format="wav")

    r = sr.Recognizer()

    # open the file
    with sr.AudioFile(r"./teste.wav") as source:
        # listen for the data (load audio to memory)
        audio_data = r.record(source, duration=5)
        # recognize (convert from speech to text)

    text2 = r.recognize_google(audio_data, language='pt-BR')
    return ''.join(text2.split(' '))


