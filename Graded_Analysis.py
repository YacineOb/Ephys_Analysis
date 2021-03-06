


'''
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
import os, time  # Operating system integration
import xlsxwriter
import pyabf  # Working with Python 3.7 (not officially on V2.0+)
import numpy as np  # Modules to manipulate arrays and some mathematical definitions
import matplotlib.pyplot as plt  # MatLab library, for plotting.
import scipy.optimize as optimize
import matplotlib.patches as patches
from pyabf import ABF  # From pyABF we want the module to read ABF files
from matplotlib import gridspec  # Matlab Layout module
from scipy.optimize import curve_fit
from scipy.signal import hilbert,find_peaks, peak_widths
from scipy import fftpack

# Introduction #########################################################################################################

start = time.time()
timestr = time.strftime("%H-%M-%S.%d.%m.%Y")             # Create file with time

# Path and file manipulations ##########################################################################################

mydirectory = 'C:/Python/InVitro/Graded'            # Change the working directory. Will be use to find
os.chdir(mydirectory)                                    # all the files we want to analyze. Change it if you want.

if not os.path.exists(mydirectory + '/Results_Graded/' + timestr):  # If no PlotsIv directory exists .format(timestr)
    os.makedirs(mydirectory + '/Results_Graded/' + timestr)  # We creat it to store our plots in.
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

    # Baseline and basic plot ###################################################
    baseline = np.mean(abf.sweepY[:PulseStart*abf.dataRate])
    max = np.amax(abf.sweepY)
    maxc = np.amax(abf.data[1])
    m = str(int(baseline)) + 'mV'
    s = max*0.10 # I use percentage because the scale can change from file to file.

    fig = plt.figure(figsize=(8, 4))
    gs = gridspec.GridSpec(2, 1, height_ratios=[20, 1])
    ax0 = plt.subplot(gs[0])
    ax0.plot(abf.sweepX, abf.sweepY, 'k', LineWidth=0.3)
    ax0.plot([abf.sweepX[0], abf.sweepX[-1]], [baseline, baseline], 'k--', LineWidth=0.5, dashes=(5, 5))
    ax0.plot([abf.sweepX[-len(abf.sweepX) // 10], abf.sweepX[-1]], [s, s], 'k', LineWidth=3,zorder=10)
    ax0.plot([abf.sweepX[-len(abf.sweepX) // 10], abf.sweepX[-len(abf.sweepX) // 10]], [s, s + 20], 'k',
             LineWidth=3,zorder=10)
    ax0.annotate(m, xy=(0.1, -54), xytext=(0, baseline + 5), fontsize=12,zorder=4)
    ax0.annotate('20 mV', xy=(0, 0), xytext=(abf.sweepX[-len(abf.sweepX) // 11], s + 10), fontsize=11,
                 fontweight='bold',zorder=4)
    ax0.annotate('10 s', xy=(0, 0), xytext=(abf.sweepX[-len(abf.sweepX) // 14], s + 2),
                 fontsize=11, fontweight='bold',zorder=4)
    ax0.add_patch(patches.Rectangle((abf.sweepX[-len(abf.sweepX) // 9 ], s - 2),10,25,
                                    fill=True,facecolor='w',zorder=3,alpha= 0.9 ))
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
    #fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Instantaneous_frequency_sweep_' + str(index + fthreshold) + '.png', dpi=400)
    plt.show()

    # Plotting Instantaneous frequency ###########################################################################
    InstantaneousFrequencyPulse = np.zeros(len(peaks) - 1)
    for pic in range(0, len(peaks) - 1):
        InstantaneousFrequencyPulse[pic] = 1 / (abf.sweepX[peaks[pic + 1]] - abf.sweepX[peaks[pic]])

    # Fit an exponential
    p0 = [100, 5, -100]
    popt, pcov = optimize.curve_fit(fitting, range(0, len(peaks) - 1), InstantaneousFrequencyPulse, p0)
    #print("a =", popt[0], "+/-", pcov[0, 0] ** 0.5)
    #print("b =", popt[1], "+/-", pcov[1, 1] ** 0.5)
    #print("c =", popt[2], "+/-", pcov[2, 2] ** 0.5)
    yEXP = fitting(range(0, len(peaks) - 1), *popt)

    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.spines['left'].set_linewidth(2)
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    ax.spines['bottom'].set_color('none')
    #plt.gca().get_xaxis().set_visible(False)  # hide X axis
    plt.scatter(range(0, len(peaks) - 1), InstantaneousFrequencyPulse, label='Data',          # !!! interesting!
            facecolors='None', edgecolors='k')
    plt.plot(range(0, len(peaks) - 1), yEXP, 'r-', ls='--', label='a=%5.1f, b=%5.1f, c=%5.1f' % tuple(popt))
    plt.text(0.55, 0.75, r'$f(t)=a{e}^{-b{t}}-c$', fontsize=20,
         transform=plt.gca().transAxes)  # Disable/Delete if you don't use Tex.
    plt.title("Instantaneous frequency - " + str(filename))
    plt.ylabel('Instantaneous frequency (Hz)', fontweight='bold')
    plt.legend()
    # fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Instantaneous_frequency_sweep_' + str(index + fthreshold) + '.png', dpi=400)
    plt.show()


    # Plotting Inter-spike Interval (ISI), fitting and Adaptation Ratio ################################################################

    x1 = range(0, len(peaks) - 1)
    y1 = InstantaneousFrequencyPulse

    trialX = np.linspace(x1[0], x1[-1], 34)

    # Fit an exponential
    popt, pcov = optimize.curve_fit(fitting, x1, y1)
    print("a =", popt[0], "+/-", pcov[0, 0] ** 0.5)
    print("b =", popt[1], "+/-", pcov[1, 1] ** 0.5)
    print("c =", popt[2], "+/-", pcov[2, 2] ** 0.5)
    yEXP = fitting(trialX, *popt)

    plt.figure()
    plt.bar(x1, y1, label='Data', color='k')
    plt.plot(trialX, yEXP, 'r-', ls='--', label="Exp Fit")
    plt.legend()
    plt.show()

    # dV/dt = f(V) ##############################################################
    window = int(Spike_Lambda * abf.dataRate)
    Average_dV = np.zeros(2*window)
    Average_mV = np.zeros(2*window)
    fAHP, DAP, sAHP  = np.zeros(len(peaks)), np.zeros(len(peaks)), np.zeros(len(peaks)) ############################
    #DAP = np.zeros(len(peaks))
    #sAHP = np.zeros(len(peaks))
    n = 0
    for spike in peaks:
        Bap, Eap = int(spike - window), int(spike + window)
        dvdt = first_derivative[Bap:Eap]
        Average_dV = [(x + y)/2 for x, y in zip(Average_dV,dvdt)]
        Average_mV = [(x + y) / 2 for x, y in zip(Average_mV, abf.sweepY[Bap:Eap])]
        plt.plot(abf.sweepY[Bap:Eap],dvdt,color='grey',LineWidth = 0.5)

        fAHP[n] = np.amin(abf.sweepY[spike:int(spike + 0.0025*abf.dataRate)])
        DAP[n] = np.amax(abf.sweepY[int(spike + 0.0025*abf.dataRate):int(spike + 0.0105*abf.dataRate)])
        sAHP[n] = np.mean(abf.sweepY[int(spike + 0.0105*abf.dataRate):int(spike + 0.0155*abf.dataRate)])
        n = n + 1

    # Average_firing = [x / len(peaks) for x in Average_firing]
    plt.plot(Average_mV,Average_dV,'r--',label = 'Average')
    plt.gca().spines['right'].set_visible(False) # Hide the right and top spines
    plt.gca().spines['top'].set_visible(False)
    plt.title('dV/dt = f(Vmb), Action Potential dynamic')
    plt.xlabel('Membrane Potential (mV)')
    plt.ylabel('dV/dt (mV/ms)')
    plt.legend()
    plt.show()

    # Plotting Threshold for ALL spikes #################################################
    plt.plot(abf.sweepX[PosPeaks], abf.sweepY[PosPeaks],'k',LineWidth=0.5)
    plt.plot([abf.sweepX[PosPeaks][0],abf.sweepX[PosPeaks][-1]],[AverageThreshold,AverageThreshold],'r--', LineWidth=1, dashes=(5, 5),label = 'Average')
    plt.gca().spines['right'].set_visible(False)
    plt.gca().spines['top'].set_visible(False)
    plt.xlabel('pulse periode (s)')
    plt.ylabel('Threshold (mV)')
    plt.legend()
    plt.title('Threshold over time during Cur2s')
    plt.show()

    # Plotting Amplitude ################################################################
    plt.plot(abf.sweepX[peaks], Amppeaks,'k',label = 'Amplitude',LineWidth=0.5)
    plt.plot([abf.sweepX[PosPeaks][0], abf.sweepX[PosPeaks][-1]], [np.mean(Amppeaks), np.mean(Amppeaks)],
             'r--',
             LineWidth=1, dashes=(5, 5), label='Average')
    # Hide the right and top spines
    plt.gca().spines['right'].set_visible(False)
    plt.gca().spines['top'].set_visible(False)
    plt.xlabel('pulse periode (s)')
    plt.ylabel('Amplitude (mV)')
    plt.legend()
    plt.title('Amplitude over Time')
    plt.show()

    print(np.mean(Amppeaks),'Mean Amplitude (mV)')

    # Width #############################################################################
    width_half = peak_widths(abf.sweepY, peaks, rel_height=1/2) # rel_height must be calculated using the peak and the threshold to find the middle.
    width_third = peak_widths(abf.sweepY, peaks, rel_height=2/3)

    plt.plot(abf.sweepX[peaks], width_half[0],'k',label = '1/2 Width')
    plt.plot(abf.sweepX[peaks], width_third[0], 'grey',label = '1/3 Width')
    plt.plot([abf.sweepX[PosPeaks][0], abf.sweepX[PosPeaks][-1]], [np.mean(width_half[0]), np.mean(width_half[0])], 'r--',
             LineWidth=1, dashes=(5, 5), label='Average')
    plt.plot([abf.sweepX[PosPeaks][0], abf.sweepX[PosPeaks][-1]], [np.mean(width_third[0]), np.mean(width_third[0])], 'r--',
             LineWidth=1, dashes=(5, 5))
    # Hide the right and top spines
    plt.gca().spines['right'].set_visible(False)
    plt.gca().spines['top'].set_visible(False)
    plt.xlabel('pulse periode (s)')
    plt.ylabel('Width (s)')
    plt.legend()
    plt.title('Width over Time')
    plt.show()

    print(np.mean(width_half[0]), 'Mean Half_width')
    print(np.mean(width_third[0]), 'Mean Third_width')

    plt.hlines(*width_half[1:], color="C2")  # That part is just to check that the width is well placed on the Y-axis
    plt.hlines(*width_third[1:], color="C1")
    plt.plot(abf.sweepY)
    plt.axis([5.0*abf.dataRate, 5.1*abf.dataRate, np.amin(abf.sweepY), np.amax(abf.sweepY) + 1])
    plt.show()

    # Fast AfterHyperpolarization (fAHP) ##########################################################

    fAHPAmplitude = np.zeros(len(peaks))
    fAHPAmplitude = [(x - y) for x, y in zip(abf.sweepY[PosPeaks],fAHP)]
    plt.plot(abf.sweepX[peaks],abf.sweepY[peaks],'xb')
    plt.plot(abf.sweepX,abf.sweepY)
    plt.plot(abf.sweepX[peaks],fAHP,'xr')
    plt.plot(abf.sweepX[peaks], DAP, 'xk')
    plt.plot(abf.sweepX[peaks], sAHP, 'xg')
    plt.axis([5.0, 5.1, np.amin(abf.sweepY), np.amax(abf.sweepY) + 1])
    plt.show()

    # Depolarizing AfterPotential (ADP) ################################################

    fAHPAmplitude = np.zeros(len(peaks))
    fAHPAmplitude = [(x - y) for x, y in zip(abf.sweepY[PosPeaks],fAHP)]
    plt.plot(abf.sweepX[peaks],fAHPAmplitude)

    plt.show()
    print(fAHPAmplitude)
