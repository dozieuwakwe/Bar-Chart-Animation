import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import ctypes
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from matplotlib.animation import FuncAnimation

#Data input
def fileimport():
    global bcrdf, importlabel, b1
    while True:
        filename = filedialog.askopenfilename()
        if 'xls' not in filename:
            ctypes.windll.user32.MessageBoxW(0, 'Please import an Excel file', "Import Error", 16)
        else:
            break
    bcrdf = pd.read_excel(filename)
    importlabel.configure(text='File imported!')
    b1.configure(text='Change Data File')

intro='Welcome to the Bar Chart Animator! Please make sure your input data is in the form of an Excel file containing a table in which the rows represent each unit of time and the columns represent each category tracked.'
ctypes.windll.user32.MessageBoxW(0, intro, "Welcome!", 64)

top = Tk(className=' Bar Chart Animator')
importlabel=Label(top, text='')
importlabel.grid(row=1, column=1)
Label(top, text='Please import a data file and fill out the information below').grid(row=0,columnspan=2)
Label(top, text='Chart Title: ').grid(row=2)
Label(top, text='X-Axis Label: ').grid(row=3)
Label(top, text='Y-Axis Label: ').grid(row=4)
Label(top, text='Time Unit: ').grid(row=5)
Label(top, text='Timesteps: ').grid(row=6)

e1=Entry(top)
e2=Entry(top)
e3=Entry(top)
e4=Entry(top)
e5=Entry(top)

e1.grid(row=2, column=1)
e2.grid(row=3, column=1)
e3.grid(row=4, column=1)
e4.grid(row=5, column=1)
e5.grid(row=6, column=1)

play=False
bcrdf=pd.DataFrame()

def callback(self):
    global title, xlabel, ylabel, units, timesteps, play
    if bcrdf.empty:
        ctypes.windll.user32.MessageBoxW(0, 'Please import a data file', "Welcome!", 16)
    else:
        title=e1.get() 
        xlabel=e2.get()
        ylabel=e3.get()
        units=e4.get()
        if e5.get()=='':
            timesteps=1
        else:
            timesteps=int(e5.get())
        top.destroy()
        play=True

top.bind('<Return>', callback)

def quit():
    top.destroy()

b1 = Button(top, text = "Import Data File", width=15, command=fileimport)
b1.grid(row=1,column=0)
b2 = Button(top, text = "Enter Information", width=15)
b2.bind('<Button-1>', callback)
b2.grid(row=7,columnspan=2)

mainloop()

#Bar chart animation
def prepare_data(df, steps):
    index=df.columns[0]
    df.index = df.index * steps
    last_idx = df.index[-1] + 1
    df_expanded = df.reindex(range(last_idx))
    df_expanded[index] = df_expanded[index].fillna(method='ffill')
    df_expanded = df_expanded.set_index(index)
    df_rank_expanded = df_expanded.rank(axis=1, method='first')
    df_expanded = df_expanded.interpolate()
    df_rank_expanded = df_rank_expanded.interpolate()
    return df_expanded, df_rank_expanded

def update(i):
    for bar in ax.containers:
        bar.remove()
    colors = plt.cm.tab20(range(20))
    y = bcrdf_rank_expanded.iloc[i]
    width = bcrdf_expanded.iloc[i]
    labels=bcrdf_expanded.columns
    ax.barh(y=y, width=width, color=colors, tick_label=labels)
    rnd = int(bcrdf_expanded.index[i])
    ax.set_ybound(lower=len(y)-19.5,upper=len(y)+0.5)
    ax.set_xbound(lower=0,upper=None)
    ax.grid(True, axis='x', color='grey')
    ax.set_axisbelow(True)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    
    if units!= '':
        heading=title+f'\n{units}: {rnd}'
    else:
        heading=title
    ax.set_title(heading, fontsize=20)
    
if play==True:
    bcrdf_expanded, bcrdf_rank_expanded = prepare_data(bcrdf,timesteps)
    fig = plt.Figure(figsize=(12, 8))
    ax = fig.add_subplot()
    anim = FuncAnimation(fig=fig, func=update, frames=len(bcrdf_expanded), interval=100, repeat=False)
    anim.save('Bar Chart Animation.mp4')
    ctypes.windll.user32.MessageBoxW(0, 'The Bar Chart animation has been successfully completed and saved!', "Success!", 64)

    #Playback/Output
    class Video(object):
        def __init__(self,path):
            self.path = path

        def play(self):
            from os import startfile
            startfile(self.path)

    class Movie_MP4(Video):
        type = "MP4"

    movie = Movie_MP4("Bar Chart Animation.mp4")
    movie.play()
