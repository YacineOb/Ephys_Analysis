'''Quick code to analyze our Cur2s files. I mostly used my MEC LII recordings to make that code. In there, the firing patterns are
not complexe and I rarely see bursts. In the future, we'll use ABF2.0 and get access to more data (such as the protocols)
that were not stored in our WinWCP files. Those new files will help us to automatize everything and avoid us to fit our code
with arbitrary values. For example, the sampling rate is already written in the ABF file Header as 'dataRate'. Coding using
that variable, allow us to share and adapt that code to every data we could have in the future.
If you have any suggestions to improve that code, the plotting, additional features, please tell me.'''

'''What that code is doing:
- Calculates the Resting Membrane Potential (RMP).
- Calculates the 1st, 2nd, 3rd and 4th derivative to find the threshold(mV).
- Calculates the Action Potential (AP) peak Rising and Falling slopes(mV/ms).
- Calculates the AP Amplitudes.
- Calculates the AP widths (at 1/2 Amplitude or 1/3 Amplitude)
- Calculates the AP ISI and ISI Ratio.
- Calculates the Fast AfterHyperPolarization (fAHP).
- Calculates the ADP.
- Calculates the sAHP.

Coming:
- Calculates the Frequency.
- Calculates the plateau potential.
- - - - - - - --  -SMPO!
- Action potential vs Threshold dynamic?
- Write everything in Excel.


from scipy import fftpack

X = fftpack.fft(x)
freqs = fftpack.fftfreq(len(x)) * f_s

fig, ax = plt.subplots()

ax.stem(freqs, np.abs(X))
ax.set_xlabel('Frequency in Hertz [Hz]')
ax.set_ylabel('Frequency Domain (Spectrum) Magnitude')
ax.set_xlim(-f_s / 2, f_s / 2)
ax.set_ylim(-5, 110)
'''

# Import modules #######################################################################################################
import pywt
import pywt.data
from scipy.interpolate import griddata

import os, time  # Operating system integration
import xlsxwriter
import pyabf  # Working with Python 3.7 (not officially on V2.0+)
import numpy as np  # Modules to manipulate arrays and some mathematical definitions
import matplotlib.pyplot as plt  # MatLab library, for plotting.
import scipy.optimize as optimize
import matplotlib.patches as patches
from pyabf import ABF  # From pyABF we want the module to read ABF files
from matplotlib import gridspec  # Matlab Layout module
from scipy import signal
from scipy.optimize import curve_fit
from scipy.signal import hilbert,find_peaks, peak_widths
from scipy import fftpack
import scipy.fftpack

# Introduction #########################################################################################################

start = time.time()
timestr = time.strftime("%H-%M-%S.%d.%m.%Y")             # Create file with time

# Path and file manipulations ##########################################################################################

mydirectory = 'C:/Python/InVitro/Analysis_Cur2s'            # Change the working directory. Will be use to find
os.chdir(mydirectory)                                    # all the files we want to analyze. Change it if you want.

if not os.path.exists(mydirectory + '/Results_Cur2s/' + timestr):  # If no PlotsIv directory exists .format(timestr)
    os.makedirs(mydirectory + '/Results_Cur2s/' + timestr)  # We creat it to store our plots in.
    print("Results path has been created")  # display message if created.
else:  # Otherwise, if the directory exists,
    pass  # Just move on.

# src_files = os.listdir(mydirectory)               # List all files in 'mydirectory'
src_files = [i for i in os.listdir(mydirectory)     # If it finds a directory, will ignore it.
             if not os.path.isdir(i)]

# Definitions #################################################################################################
def fitting(t, a, b, c):                               # Fitting equation. Choose what is best.
    return a * np.exp(-b * t) - c


# Variable definition (You can change that according to your protocols)
PulseStart = 5
PulseEnd = 7
Spike_Lambda = 0.02

Rslt = {'Filename': 0, 'Resistance Average (MOhm)': 0, 'Baseline Average (mV)': 0, 'Baseline Average before Onset (mV)': 0,
        'Steady state depolarisation (mV)': 0, 'Sag Value (mV)': 0, 'Sag Ratio (mV)': 0, 'Sag Peak (mV)': 0}

# Excel file opening ###################################################################################################
workbook = xlsxwriter.Workbook(mydirectory + '/Results_Cur2s/' + timestr + '/Results_' + timestr + '.xlsx')
worksheet = workbook.add_worksheet('Basic_properties')

col = 0
row = 0
for key in Rslt.keys():
    worksheet.write(row, col, key) # row, col, item
    col += 1

# ABF file opening #####################################################################################################
for filename in src_files:
    os.makedirs(mydirectory + '/Results_Cur2s/' + timestr + '/' + filename[:-4])
    abf: ABF = pyabf.ABF(filename)

    # ABF-dependant Variables definitions - Extracting basic properties ################################################
    # Peaks during pulse - Main process on data ########################################################################
    peaks, _ = find_peaks(abf.sweepY[PulseStart*abf.dataRate:
                                     PulseEnd*abf.dataRate], height=-10, distance=10)

    '''
    # Derivatives ######################################################################################################
    first_derivative = np.gradient(abf.sweepY)
    second_derivative = np.gradient(first_derivative)
    third_derivative = np.gradient(second_derivative)

    # AP Slopes for all spikes #################################################
    pospeaksSlope, _ = find_peaks(first_derivative)
    negpeaksSlope, _ = find_peaks(-first_derivative)

    meanpeaksplope = np.mean(first_derivative[pospeaksSlope])
    meannegslope = np.mean(first_derivative[negpeaksSlope])

    print(meanpeaksplope,'Mean rising slope (mV/ms)',meannegslope,'Mean falling slope (mV/ms)')

    # Threshold for ALL spikes #################################################
    PosPeaks, _ = find_peaks(third_derivative, height=1.5) # find the height for next with threshold
    AverageThreshold = np.mean(abf.sweepY[PosPeaks])

    # Amplitude ################################################################
    Amppeaks = [(x - y) for x, y in zip(abf.sweepY[peaks],abf.sweepY[PosPeaks])]

    # Frequency ########################################################################################################
    Frequency = len(peaks)

    # Instantaneous frequency ##########################################################################################
    InstantaneousFrequencyPulse = np.zeros(len(peaks) - 1)
    for pic in range(0, len(peaks) - 1):
        InstantaneousFrequencyPulse[pic] = 1 / (abf.sweepX[peaks[pic + 1]] - abf.sweepX[peaks[pic]])

    # Inter-spike Interval (ISI) and Adaptation Ratio ############################################
    FirstISI = (abf.sweepX[peaks[1]] - abf.sweepX[peaks[0]])
    LastISI = (abf.sweepX[peaks[-1]] - abf.sweepX[peaks[-2]])
    print(FirstISI, 'First ISI')
    print(LastISI, 'Last ISI')
    print(LastISI / FirstISI, 'Adaptation Ratio')
    '''
    # Baseline and basic plot ###################################################
    baseline = np.mean(abf.sweepY[:PulseStart*abf.dataRate])
    max = np.amax(abf.sweepY)
    maxc = np.amax(abf.data[1])
    m = str(int(baseline)) + 'mV'
    s = max*0.10 # I use percentage because the scale can change from file to file.

    fig = plt.figure(figsize=(8, 4))
    gs = gridspec.GridSpec(3, 1, height_ratios=[20, 1, 2])
    ax0 = plt.subplot(gs[0])
    ax0.plot(abf.sweepX, abf.sweepY, 'k', LineWidth=0.3)
    ax0.plot([abf.sweepX[0], abf.sweepX[-1]], [baseline, baseline], 'k--', LineWidth=0.5, dashes=(5, 5))
    # ax0.plot([abf.sweepX[-len(abf.sweepX) // 10], abf.sweepX[-1]], [s, s], 'k', LineWidth=3,zorder=10)
    # ax0.plot([abf.sweepX[-len(abf.sweepX) // 10], abf.sweepX[-len(abf.sweepX) // 10]], [s, s + 20], 'k',
    #          LineWidth=3,zorder=10)
    ax0.annotate(m, xy=(0.1, -54), xytext=(0, baseline + 5), fontsize=12,zorder=4)
    # ax0.annotate('20 mV', xy=(0, 0), xytext=(abf.sweepX[-len(abf.sweepX) // 11], s + 10), fontsize=11,
    #               fontweight='bold',zorder=4)
    # ax0.annotate('10 s', xy=(0, 0), xytext=(abf.sweepX[-len(abf.sweepX) // 14], s + 2),
    #              fontsize=11, fontweight='bold',zorder=4)
    # ax0.add_patch(patches.Rectangle((abf.sweepX[-len(abf.sweepX) // 9 ], s - 2),10,25,
    #                                 fill=True,facecolor='w',zorder=3,alpha= 0.9 ))
    for spine in plt.gca().spines.values():
        spine.set_visible(False)
    plt.gca().get_yaxis().set_visible(False)  # hide Y axis
    plt.gca().get_xaxis().set_visible(False)  # hide Y axis
    plt.tight_layout()

    ax1 = plt.subplot(gs[1])
    for x in abf.sweepList:
        ax1.plot(abf.sweepX, abf.data[1][len(abf.sweepX) * x:len(abf.sweepX) * (x + 1)], 'k', LineWidth=1)
    for spine in plt.gca().spines.values():
        spine.set_visible(False)
    #ax1.plot([abf.sweepX[-len(abf.sweepX) // 10], abf.sweepX[-len(abf.sweepX) // 10]], [abf.data[1][0], maxc], 'k',
    #         LineWidth=3)
    #ax1.annotate('200 pA', xy=(0, 0), xytext=(abf.sweepX[-len(abf.sweepX) // 11], maxc / 2), fontsize=2,
    #             fontweight='bold')
    plt.gca().get_yaxis().set_visible(False)  # hide Y axis
    plt.gca().get_xaxis().set_visible(False)  # hide Y axis
    plt.tight_layout()
    ###########################################################################################
    #ax2 = plt.subplot(gs[2])
    #plt.bar(InstantaneousFrequencyPulseb peaks[:-1], )

    '''
    ###########################################################################################
    InstantaneousFrequency = [1 / (item[j + 1] - item[j]) * abf.dataRate for j in np.arange(0, len(item) - 1)]

        # Fitting Instantaneous frequency process
    p0 = [1000, 5, -100] # The function need some help on the parameters. input your first guess here.
    popt, pcov = optimize.curve_fit(fitting, range(0, len(PeaksSweep[index + fthreshold]) - 1),
                                        InstantaneousFrequency[index + fthreshold], p0, maxfev=5000)
        # print("a =", popt[0], "+/-", pcov[0, 0] ** 0.5)
        # print("b =", popt[1], "+/-", pcov[1, 1] ** 0.5)
        # print("c =", popt[2], "+/-", pcov[2, 2] ** 0.5)
    yEXP = fitting(np.arange(0, len(PeaksSweep[index + fthreshold]) - 1), *popt)

        # Plotting processed data ######################################################################################

        # Instantaneous frequency and fitting plot
    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.spines['left'].set_linewidth(2)
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    ax.spines['bottom'].set_color('none')
    plt.gca().get_xaxis().set_visible(False)  # hide X axis
    plt.scatter(range(0, len(PeaksSweep[index + fthreshold]) - 1), InstantaneousFrequency[index + fthreshold], label='Data', facecolors='None', edgecolors='k')
    plt.plot(range(0, len(PeaksSweep[index + fthreshold]) - 1), yEXP, 'r-', ls='--', label='a=%5.1f, b=%5.1f, c=%5.1f' % tuple(popt))
    plt.text(0.55, 0.75, r'$f(t)=a{e}^{-b{t}}-c$', fontsize=20,
            transform=plt.gca().transAxes)  # Disable/Delete if you don't use Tex.
    plt.title("Instantaneous frequency - " + str(filename) + "; \n $n_{sweep} =$ " + str(index + fthreshold))
    plt.ylabel('Instantaneous frequency (Hz)', fontweight='bold')
    plt.legend()
    #fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Instantaneous_frequency_sweep_' + str(index + fthreshold) + '.png', dpi=400)
    #plt.show()

    '''
    ###########################################################################################
    fig.savefig(mydirectory + '/Results_Cur2s/' + timestr + '/' + filename[:-4] + '/' + '-' + '.png', dpi=400)
    plt.show()
