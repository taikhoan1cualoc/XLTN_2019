import tkinter as tk
import pickle as cPickle
import numpy as np
from scipy.io.wavfile import read
from sklearn import mixture
from sklearn.mixture import gaussian_mixture as GMM
from featureExtraction import extract_features
from tkinter import ttk
import warnings
import pyaudio
import wave
import os
import webbrowser
import ctypes
import wmi
from win32com.client import GetObject


CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 2
RATE = 44100
RECORD_SECONDS = 2

def execute(cmd):
    brightness = None
    if cmd == 'chrome':
        webbrowser.get("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s").open("http://google.com.vn")
    elif cmd == 'khoa may':
        ctypes.windll.user32.LockWorkStation()
    elif cmd == 'tang do sang':
        objWMI = GetObject('winmgmts:\\\\.\\root\\WMI').InstancesOf('WmiMonitorBrightness')
        for obj in objWMI:
            if obj.CurrentBrightness != None:
                brightness = obj.CurrentBrightness # percentage [0-100]
                break
       # print brightness
        if brightness <= 90:
            brightness += 10
        c = wmi.WMI(namespace='wmi')
        methods = c.WmiMonitorBrightnessMethods()[0]
        methods.WmiSetBrightness(brightness, 0)
    elif cmd == 'giam do sang':
        objWMI = GetObject('winmgmts:\\\\.\\root\\WMI').InstancesOf('WmiMonitorBrightness')
        for obj in objWMI:
            if obj.CurrentBrightness != None:
                brightness = obj.CurrentBrightness  # percentage [0-100]
                break
        if brightness >= 10:
            brightness -= 10
        c = wmi.WMI(namespace='wmi')
        methods = c.WmiMonitorBrightnessMethods()[0]
        methods.WmiSetBrightness(brightness, 0)

def test():
    display = tk.Text(master=window, height=8, width=40)
    display.config(font=("Helvetica"))
    display.grid(columnspan=2, row=5, sticky='e')
    WAVE_OUTPUT_FILENAME = "test/test"

    # #path to training data
    source = "data/"
    modelpath = "models/"
    # test_file = "development_set_test.txt"
    # file_paths = open(test_file,'r')
    #

    gmm_files = [os.path.join(modelpath, fname) for fname in
                 os.listdir(modelpath) if fname.endswith('.gmm')]
    # Load the Gaussian gender Models
    models = [cPickle.load(open(fname, 'rb')) for fname in gmm_files]
    #models = [cPickle.load(open("giam do sang.gmm", 'rb')) ]

    speakers = [fname.split("/")[-1].split(".gmm")[0] for fname
                in gmm_files]

    # Read the test directory and get the list of test audio files
    p = pyaudio.PyAudio()
    stream = p.open(format=FORMAT,
                    channels=CHANNELS,
                    rate=RATE,
                    input=True,
                    frames_per_buffer=CHUNK)

    # print("* recording")

    frames = []

    for i in range(0, int(RATE / CHUNK * RECORD_SECONDS)):
        data = stream.read(CHUNK)
        frames.append(data)

    # print("* done recording")
    stream.stop_stream()
    stream.close()
    p.terminate()
    wf = wave.open(WAVE_OUTPUT_FILENAME + ".wav", 'wb')

    wf.setnchannels(CHANNELS)
    wf.setsampwidth(p.get_sample_size(FORMAT))
    wf.setframerate(RATE)
    wf.writeframes(b''.join(frames))
    wf.close()

    sr, audio = read("test/test.wav")
    vector = extract_features(audio, sr)

    log_likelihood = np.zeros(len(models))

    for i in range(len(models)):
        gmm = models[i]  # checking with each model one by one
        scores = np.array(gmm.score(vector))
        log_likelihood[i] = scores.sum()

    winner = np.argmax(log_likelihood)
    # print "\tDETECTED AS: ", speakers[winner]
    display.insert(tk.END, "DETECTED AS: " + speakers[winner])
    execute(speakers[winner])

def start_record():
    display = tk.Text(master=window, height=8, width=40)
    display.config(font=("Helvetica"))
    display.grid(columnspan=2, row=5, sticky='e')
    input = entry.get()
    directory = "data/" + str(input) + "/"
    if not os.path.exists(directory):
        os.makedirs(directory)
    count = 0
    while os.path.exists(directory + "w" + str(count)+".wav"):
        count += 1
    p = pyaudio.PyAudio()
    stream = p.open(format=FORMAT,
                    channels=CHANNELS,
                    rate=RATE,
                    input=True,
                    frames_per_buffer=CHUNK)

    # print("* recording")
    display.insert(tk.END, "")
    frames = []

    for i in range(0, int(RATE / CHUNK * RECORD_SECONDS)):
        data = stream.read(CHUNK)
        frames.append(data)

    # print("* done recording")
    display.insert(tk.END, "* Done recording: " + input + "/w" + str(count) + ".wav")
    display.insert(tk.END, "\n")
    stream.stop_stream()
    stream.close()
    p.terminate()
    wf = wave.open(directory + "w" + str(count)+".wav", 'wb')
    wf.setnchannels(CHANNELS)
    wf.setsampwidth(p.get_sample_size(FORMAT))
    wf.setframerate(RATE)
    wf.writeframes(b''.join(frames))
    wf.close()
    with open("data.txt", "a") as file:
        file.write(input + "/" + "w" + str(count)+".wav")
        file.write("\n")
        file.close()

def train():
    display = tk.Text(master=window, height=8, width=40)
    display.config(font=("Helvetica"))
    display.grid(columnspan=2, row=5, sticky='e')
    warnings.filterwarnings("ignore")

    # path to training data
    source = "data/"

    # path where training speakers will be saved
    dest = "models/"
    train_file = "data.txt"
    file_paths = open(train_file, 'r')

    count = 1
    # Extracting features for each speaker (5 files per speakers)
    features = np.asarray(())
    for path in file_paths:
        path = path.strip()
        # print path
        # display.insert(tk.END, path)
        # display.insert(tk.END, "\n")

        # read the audio
        sr, audio = read(source + path)

        # extract 40 dimensional MFCC & delta MFCC features
        vector = extract_features(audio, sr)

        if features.size == 0:
            features = vector
        else:
            features = np.vstack((features, vector))
        # when features of 5 files of speaker are concatenated, then do model training
        if count == 5:
            gmm = GMM(n_components=16, n_iter=200, covariance_type='diag', n_init=3)
            gmm.fit(features)

            # dumping the trained gaussian model
            picklefile = path.split("/")[0] + ".gmm"
            # print picklefile
            # f = open(dest + picklefile, 'w+')
            cPickle.dump(gmm, open(dest + picklefile, 'w+'))
            # print '+ modeling completed for word:', picklefile, " with data point = ", features.shape
            phrase = "Modeling completed: " + picklefile
            display.insert(tk.END, phrase)
            display.insert(tk.END, "\n")
            features = np.asarray(())
            count = 0
        count = count + 1

window = tk.Tk()
window.style = ttk.Style()
window.style.theme_use("default")
window.title("Voice Command")
window.geometry("380x300")
title = tk.Label(text="Voice Command")
title.config(font=("Helvetica", 30))
title.grid(columnspan=2, row=0)
trainlb = tk.Label(text="Training Models")
trainlb.grid(column=0, row=1, sticky='e')
btnTrain = tk.Button(text="Train", width=28, command=train)
btnTrain.grid(column=1, row=1)
trainlb = tk.Label(text="Voice Recognition")
trainlb.grid(column=0, row=2, sticky='e')
btnTest = tk.Button(text="Voice", width=28, command=test)
btnTest.grid(column=1, row=2)
trainlb = tk.Label(text="New Data")
trainlb.grid(column=0, row=3, sticky='e')
entry = tk.Entry(width=28)
entry.grid(column=1, row=3)
trainlb = tk.Label(text="Record New Training Data")
trainlb.grid(column=0, row=4, sticky='e')
btnRecord = tk.Button(text="Record", width=28, command=start_record)
btnRecord.grid(column=1, row=4)
display = tk.Text(master=window, height=8, width=40)
display.grid(columnspan=2, row=5, sticky='e')
window.mainloop()
