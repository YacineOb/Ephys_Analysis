""" IV Analysis by Yacine, January 2020 - Work in Progress code.

What this code is doing:
- Calculates Baseline for ALL sweeps and the average Baseline.
- Calculates Sag peaks for ALL sweeps if possible.
- Calculates Sag Amplitude for ALL sweeps if possible.
- Calculates Sag Ratio for ALL sweeps if possible.
- Calculates Sag Exp. fitting + Tau constant for first sweep.
- Calculates the firing frequency for ALL sweeps when Vmb > Vrest.
- Calculates the instantaneous firing frequency and Exp. fit for ALL sweeps when Vmb > Vrest.
- Calculates The Resistance for each sweep and average resistance.
- Calculates the ISI Adaptation

What is coming:
- Calculates the Resistance pre vs inter.
"""

# Import modules #######################################################################################################
import os, time  # Operating system integration
import xlsxwriter
import pyabf  # Working with Python 3.7 (not officially on V2.0+)
import numpy as np  # Modules to manipulate arrays and some mathematical definitions
import matplotlib.pyplot as plt  # MatLab library, for plotting.
import scipy.optimize as optimize
from pyabf import ABF  # From pyABF we want the module to read ABF files
from matplotlib import gridspec  # Matlab Layout module
from scipy.optimize import curve_fit
from scipy.signal import find_peaks, peak_widths
import scipy.signal
import warnings
import math

# Introduction #########################################################################################################

start = time.time()
timestr = time.strftime("%H-%M-%S.%d.%m.%Y")             # Create file with time

# Path and file manipulations ##########################################################################################

mydirectory = 'C:/Python/InVitro/Analysis_IV'            # Change the working directory. Will be use to find
os.chdir(mydirectory)                                    # all the files we want to analyze. Change it if you want.

if not os.path.exists(mydirectory + '/Results_IV/' + timestr):  # If no PlotsIv directory exists
    os.makedirs(mydirectory + '/Results_IV/' + timestr)         # We creat it to store our plots in.
    print("Results path has been created")                      # display message if created.
else:                                                           # Otherwise, if the directory exists,
    pass                                                        # Just move on.

# src_files = os.listdir(mydirectory)                           # List all files in 'mydirectory'
src_files = [i for i in os.listdir(mydirectory)                 # If it finds a directory, will ignore it.
             if not os.path.isdir(i)]

# Definitions #################################################################################################
def fitting(t, a, b, c):                               # Fitting equation - CHECK THE EQUATION!
    return a * np.exp(-b * t) - c

#def Resistance() - coming soon.

''' More efficient way to store our values. Coming soon.
Cstsec = {'PulseStart':1.50, 'PulseEnd': 2.50, 'ResistancePulseStart': 4.00,'RPulseInterval': 0.500,
       'ResistancePulseEnd': 4.50, 'SagFrame': 0.150,'AverageOffOnSet': 0.200, 'RPulseInterval': 0.500}             

Var = {'Vsteady': 0,'VHcurrent':0, 'ResistanceSweep':0, 'ResistanceSweep':0, 'BaselineSweep':0,
            'VBaseline':0, 'PeaksSweep':0, 'Frequency':0, 'InstantaneousFrequency':0}
'''

Rslt = {'Filename': 0, 'Resistance Average (MOhm)': 0,'Resistance Sag (MOhm)':0,'Resistance Steady-state (MOhm)':0, 'Baseline Average (mV)': 0, 'Baseline Average before Onset (mV)': 0,
        'Steady state depolarisation (mV)': 0, 'Sag Max Value (mV)': 0,'Sag Amplitude (mV)':0, 'Sag Ratio': 0, 'Sag Peak (mV)': 0,
        'Sag Full-Width Half maximum (ms)':0,'SagFit_a': 0,'SagFit_b': 0,'SagFit_c': 0,}

# Excel file opening ###################################################################################################
workbook = xlsxwriter.Workbook(mydirectory + '/Results_IV/' + timestr + '/Results_' + timestr + '.xlsx')
worksheet = workbook.add_worksheet('Basic_properties')

col = 0
row = 0
for key in Rslt.keys():
    worksheet.write(row, col, key) # row, col, item
    col += 1

# ABF file opening #####################################################################################################

for filename in src_files:
    os.makedirs(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4])
    abf: ABF = pyabf.ABF(filename)

    #filename = filename[:-10]            # Name your file as you want
    #print(abf.headerText)

    print("Reading", filename + '...', end=""),

    # ABF-dependant Variables definitions ##############################################################################

    CurrentIn = np.linspace(-300, 300, len(abf.sweepList))  # ! I create the current input data. Use SweepC for ABF2.0. No SweepC with WinWCP converted ABF

    Vsteady =  np.zeros(len(abf.sweepList))
    VHcurrent = ResistanceSweep = np.zeros(len(abf.sweepList))
    ResistanceSweep = np.zeros(len(abf.sweepList))
    BaselineSweep = np.zeros(len(abf.sweepList))
    VBaseline = np.zeros(len(abf.sweepList))
    Frequency = np.zeros(len(abf.sweepList))
    PeaksSweep = np.zeros(len(abf.sweepList),dtype=object)
    InstantaneousFrequency = np.zeros(len(abf.sweepList), dtype=object)
    AdaptationRatio = np.zeros(len(abf.sweepList), dtype=object)

    # Cst = {key: int(Cst[key] * abf.dataRate) for key in Cst.keys()}   # Use for dictionnary above. Next update.

    PulseStart = int(1.65*abf.dataRate)                     # Seconds. Depends on the protocol. we'll be in a dict soon. 1.5 before. 1.65 now?!
    PulseEnd = int(2.65*abf.dataRate)
    ResistancePulseStart = int(4.65*abf.dataRate)
    RPulseInterval = int(0.500*abf.dataRate)
    ResistancePulseEnd = int(5.15*abf.dataRate)
    SagFrame = int(0.165*abf.dataRate)
    AverageOffOnSet = int(0.200*abf.dataRate)
    RPulseInterval = int(0.500*abf.dataRate)


    # Algorithm core - Main iteration ##################################################################################

    for sweepNumber in abf.sweepList:
        abf.setSweep(sweepNumber)

        # Baseline from 0 second to the beginning of the pulse.
        BaselineSweep[sweepNumber] = np.mean(abf.sweepY[:PulseStart])  # Mean/baseline before current pulse
        VBaseline[sweepNumber] = np.mean(abf.sweepY[PulseStart - AverageOffOnSet:PulseStart ])
        ResistanceSweep[sweepNumber] = (np.mean(abf.sweepY[ResistancePulseStart:ResistancePulseEnd])
                                        -np.mean(abf.sweepY[ResistancePulseStart - RPulseInterval:ResistancePulseStart]))\
                                       /(np.mean(abf.sweepC[ResistancePulseStart:ResistancePulseEnd])*1e-3)  # real values in data var.
                                                                                                             # deltaV for Resistance
        # If the mean during the pulse frame is bellow the baseline +5 mV, we calculate the sag.
        if np.mean(abf.sweepY[PulseStart:PulseEnd]) <= BaselineSweep[sweepNumber] + 5:
            VHcurrent[sweepNumber] = np.amin(abf.sweepY[PulseStart:PulseStart + SagFrame])  # Take the minimum Voltage for each sweep.
            Vsteady[sweepNumber] = np.mean(abf.sweepY[PulseEnd - AverageOffOnSet:PulseEnd])  # Time (secondes) multiplied by the data rate = time in sample point.

        # If the mean during the pulse frame is equal to the baseline, there is probably no sag and no spikes.
        elif np.mean(abf.sweepY[PulseStart:PulseEnd]) == BaselineSweep[sweepNumber] + 5:
            continue

        # If the mean during the pulse frame is above the baseline, count spikes.
        else:
            PeaksSweep[sweepNumber], _ = find_peaks(abf.sweepY, height=-5)
            Frequency[sweepNumber] = len(PeaksSweep[sweepNumber])

    print(" Processing data...", end="")

    # Basic properties layout ##########################################################################################

    if math.isnan(float(np.mean(ResistanceSweep))) or math.isinf(np.mean(ResistanceSweep)) == True:
        warnings.warn(filename + " doesn't have -5pA input. Resistance is set to 0 ")
        ResistanceAverage = 0
    else:
        ResistanceAverage = np.mean(ResistanceSweep)

    BaselineAverage = np.mean(BaselineSweep)
    SagPeak = BaselineSweep[0] - VHcurrent[0]                # From baseline. Not 'correct' name. Rename if you want.
    SagRatio = (VHcurrent[0] - Vsteady[0])/(VHcurrent[0] - VBaseline[0])


    ISIthreshold = list(map(lambda i: i >= 8, Frequency)).index(True)  # Take all sweeps ∋ nSpike > 8. For ISI adapt.
    for index, item in enumerate(PeaksSweep[ISIthreshold:]):
        AdaptationRatio[index + ISIthreshold] = [(item[-1] - item[-2]) / (item[1] - item[0])]

    fthreshold = list(map(lambda i: i > 3, Frequency)).index(True)     # Take all sweeps ∋ nSpike > 3. For fitting.
    for index, item in enumerate(PeaksSweep[fthreshold:]):
        InstantaneousFrequency[index + fthreshold] = [1 / (item[j + 1] - item[j]) * abf.dataRate for j in np.arange(0, len(item) - 1)]

        # Fitting Instantaneous frequency process
        p0 = [-100, 1, 100] # The function need some help on the parameters. input your first guess here.
        popt, pcov = optimize.curve_fit(fitting, range(0, len(PeaksSweep[index + fthreshold]) - 1),
                                        InstantaneousFrequency[index + fthreshold], p0, maxfev=5000)
        # print("a =", popt[0], "+/-", pcov[0, 0] ** 0.5)
        # print("b =", popt[1], "+/-", pcov[1, 1] ** 0.5)
        # print("c =", popt[2], "+/-", pcov[2, 2] ** 0.5)
        yEXP = fitting(np.arange(0, len(PeaksSweep[index + fthreshold]) - 1), *popt)

        # Plotting processed data ######################################################################################
        '''
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
        fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Instantaneous_frequency_sweep_' + str(index + fthreshold) + '.png', dpi=400)
        #plt.show()
        '''

    # Plotting processed data ##########################################################################################
    # Plotting the first sag trace, baseline, sag amplitude, ratio and Vsteady versus Vsag #############################
    abf.setSweep(abf.sweepList[0])  # We want to plot the first trace only


    # I-V Plot #########################################################################################################
    fig = plt.figure()
    ax = fig.add_subplot(111)

    plt.scatter(CurrentIn[:len(np.trim_zeros(Vsteady))],np.trim_zeros(Vsteady),
                label='Steady state', facecolors='white', edgecolors='k')
    plt.plot(CurrentIn[:len(np.trim_zeros(Vsteady))],scipy.signal.savgol_filter(np.trim_zeros(Vsteady), 5, 2),
             linewidth=1.5, linestyle='--', color='k', label='Steady state fitted')#, marker='o')

    plt.scatter(CurrentIn[:len(np.trim_zeros(VHcurrent))], np.trim_zeros(VHcurrent),
                label='Peak', facecolors='white', edgecolors='b')
    plt.plot(CurrentIn[:len(np.trim_zeros(VHcurrent))], scipy.signal.savgol_filter(np.trim_zeros(VHcurrent), 5, 2),
             linewidth=1.5, linestyle='--', color='k', label='Peak fitted')

    for axis in ['bottom', 'left']:
        ax.spines[axis].set_linewidth(2)
    ax.spines['right'].set_color('none')  # Eliminate upper and right axes
    ax.spines['top'].set_color('none')
    # ax.spines['left'].set_position('center')     # Move left y-axis and bottim x-axis to centre, passing through (0,0)
    # ax.spines['bottom'].set_position('center')   # Nice but not very efficient with ABF1. Lets see later with ABF2
    # ax.xaxis.set_ticks_position('bottom')        # Show ticks in the left and lower axes only. Go with the code above.
    # ax.yaxis.set_ticks_position('left')
    #plt.axhline(y=0, linewidth=1, color='k', linestyle='dotted')
    plt.title("Current-Potential relationship")
    plt.ylabel('Potential (mV)', fontweight='bold')
    plt.xlabel('Current Injected (pA)', fontweight='bold')
    plt.legend()
    fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Current-Potential_relationship.png', dpi=400)
    #plt.show()

    # Plotting the frequency of sweeps #################################################################################

    fig = plt.figure()
    ax = fig.add_subplot(111)
    for axis in ['bottom', 'left']:
        ax.spines[axis].set_linewidth(2)

    plt.scatter(CurrentIn,Frequency,
                label='Steady state', facecolors='w', edgecolors='k')
    plt.plot(CurrentIn, scipy.signal.savgol_filter(Frequency, 5, 3), linewidth=1.5,
             linestyle='--', color='k')#, marker='o')

    ax.spines['right'].set_color('none')  # Eliminate upper and right axes
    ax.spines['top'].set_color('none')
    plt.xlim(left=-50)  # adjust the left leaving right unchanged
    plt.title("Current-Frequency relationship")
    plt.xlabel('Current Injected (pA)', fontweight='bold')
    plt.ylabel('Average Firing Frequency (Hz)', fontweight='bold')
    fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Frequency.png', dpi=400)
    #plt.show()


    #FWHM######################################################################################
    y = abf.sweepY[PulseStart:PulseEnd]
    x = abf.sweepX[PulseStart:PulseEnd]
    c = abf.sweepC[PulseStart:PulseEnd]

    peaksPC, _ = find_peaks(-y, prominence=(5),distance=5000)
    results_half = peak_widths(-y, peaksPC, rel_height=0.5)

    SagHalfWidth = results_half[0] / 20000 * 10e2
    #print("half W:", SagHalfWidth,
    #      "peak:", peaksPC)

    #plt.plot(x*20000, y, 'k', label='Data')
    #plt.plot(1.65*20000+peaksPC, y[peaksPC], "rx")
    #plt.hlines(results_half[1]*-1,1.65*20000+results_half[2],1.65*20000+results_half[3], color='red')
    #plt.xlim([PulseStart,PulseEnd])
    #plt.plot(x[int(peaksPC)-100:int(peaksPC)+100]*20000,y[int(peaksPC)-100:int(peaksPC)+100])#,color='blue')
    #plt.show()

    #Resistance (MOhm)#######################################

    SagPotential = np.mean(y[int(peaksPC)-100:int(peaksPC)+100])
    InputCurrent = np.mean(c[int(peaksPC)-100:int(peaksPC)+100])
    SagResistance = (VBaseline[0]-(SagPotential))/-(InputCurrent*1e-3)
    SteadyResistance = (VBaseline[0]-Vsteady[0])/-(np.mean(abf.sweepC[PulseEnd - AverageOffOnSet:PulseEnd])*1e-3)

    #Plot sag with ratio########################

    fig = plt.figure()
    ax = fig.add_subplot(111)
    plt.plot(abf.sweepX, abf.sweepY, 'k', linewidth=1)
    plt.axis([1.25, 3, np.amin(abf.sweepY) - 2, BaselineSweep[0] + 10])  # plt.axis([xmin,xmax,ymin,ymax])
    for axis in ['bottom', 'left']:
        ax.spines[axis].set_linewidth(2)
    plt.gca().spines['right'].set_visible(False)
    plt.gca().spines['top'].set_visible(False)

    plt.annotate('', xy=(2.65, Vsteady[0]), xytext=(2.65, VHcurrent[0]), arrowprops=dict(arrowstyle='<->'))

    plt.annotate(str(round(SagRatio, 2)) + ' Sag Ratio', xy=(2.65, Vsteady[0]),
                 xytext=(2.70, Vsteady[0]+(Vsteady[0] - VHcurrent[0]) * - 0.45))  # Sag Ratio
    plt.annotate(str(round(Vsteady[0] - VHcurrent[0], 2)) + ' mV', xy=(2.70, Vsteady[0]),
                 xytext=(2.70, VHcurrent[0] + (VHcurrent[0]-Vsteady[0])*-0.25))  # Vsag - Vsteady

    plt.annotate('', xy=(1.50, VHcurrent[0]), xytext=(1.50, BaselineSweep[0]), arrowprops=dict(arrowstyle='<->'))
    plt.annotate(str(round(SagPeak, 2)) + ' mV', xy=(1.28, VHcurrent[0]),
                 xytext=(1.45, (VHcurrent[0] + BaselineSweep[0]) / 2), rotation=90, va='center')  # Sag Amplitude
    plt.annotate(str(round(np.mean(BaselineSweep), 2)) + ' mV', xy=(1.25, BaselineSweep[0]),
                 xytext=(1.50, BaselineSweep[0] + 1.25), ha='center')  # Baseline

    plt.axhline(y=VHcurrent[0], linewidth=0.7, color='r', linestyle='--')
    plt.axhline(y=Vsteady[0], linewidth=0.7, color='b', linestyle='--')
    plt.axhline(y=BaselineSweep[0], linewidth=0.7, color='k', linestyle='--')

    plt.hlines(results_half[1]*-1,(1.65*20000+results_half[2])/20000,(1.65*20000+results_half[3])/20000, color='red')

    plt.title("Sag properties")
    plt.xlabel('Time (s)', fontweight='bold')
    plt.ylabel('Potential (mV)', fontweight='bold')
    fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Sag_properties.png', dpi=400) # Disable if you don't need to plot. Increases speed.
    plt.show()


    # Sag magnification and exponential fitting ########################################################################

    IndexSag = np.amin(np.where(abf.sweepY == VHcurrent[0]))

    x1 = abf.sweepX[int(IndexSag):int(2.5 * abf.dataRate)]
    y1 = abf.sweepY[int(IndexSag):int(2.5 * abf.dataRate)]

    fig = plt.figure()
    ax = fig.add_subplot(111)
    plt.plot(x1, y1, 'grey', label='Data')

    p1 = [10000, 5, -50]  # Change if you think the dynamic is different.
    popta, pcova = optimize.curve_fit(fitting, x1, y1, p1, maxfev=10000)  # Big number of iteration.
    #print("a2 =", popta[0], "+/-", pcova[0, 0] ** 0.5)
    #print("b2 =", popta[1], "+/-", pcova[1, 1] ** 0.5)
    #print("c2 =", popta[2], "+/-", pcova[2, 2] ** 0.5)
    yEXP1 = fitting(x1, *popta)

    plt.plot(x1, yEXP1, 'r-', ls='--', label='a=%5.3f, b=%5.3f, c=%5.3f' % tuple(popt))
    for axis in ['bottom', 'left']:
        ax.spines[axis].set_linewidth(2)
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')
    plt.text(0.55, 0.2, r'$V(t)=a{e}^{-b{t}}-c$', fontsize=20,
             transform=plt.gca().transAxes)  # Disable/Delete if you don't use Tex.
    plt.xlabel('Time (s)', fontweight='bold')
    plt.ylabel('Potential (mV)', fontweight='bold')
    plt.legend()
    fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'SagFit.png', dpi=400)
    #plt.show()

    
    # Plotting a figure with the first and last sweep.###############################
    fig = plt.figure(figsize=(10, 2))
    gs = gridspec.GridSpec(2, 1, height_ratios=[10, 1])
    axs = plt.subplots(2, 1, sharex=True, gridspec_kw={'wspace': 0, 'hspace': 0})# Remove horizontal space between axes
    ax0 = plt.subplot(gs[0])# Plot each graph, and manually set the y tick values

    limitmin = np.amin(abf.sweepY)

    abf.setSweep(abf.sweepList[-1])  # Change sweep.
    limitmax = np.amax(abf.sweepY)
    ax0.plot(abf.sweepX, abf.sweepY, 'grey', linewidth=1)


    abf.setSweep(abf.sweepList[12])
    ax0.plot(abf.sweepX, abf.sweepY, 'blue', linewidth=1)

    for spine in plt.gca().spines.values():
        spine.set_visible(False)
    #ax0.annotate(str(round(np.mean(BaselineSweep), 2)) + ' mV', xy=(1.2, np.mean(BaselineSweep)),
    #             xytext=(1.2, np.mean(BaselineSweep) + 10), fontsize=10, fontweight='bold', ha='center')# change baselinesweep value with a simple variable
    plt.gca().get_yaxis().set_visible(False)  # hide Y axis
    plt.gca().get_xaxis().set_visible(False)  # hide Y axis
    plt.axis([1.5, 2.85, limitmin, limitmax])  # plt.axis([xmin,xmax,ymin,ymax])
    #plt.plot(Vsteady[0], Vsteady[0], abf.data[1][0], abf.data[1][-1], 'r')

    ax1 = plt.subplot(gs[1])
    ax1.plot(abf.sweepX, abf.data[1][len(abf.sweepX) * 8:len(abf.sweepX) * 9], 'b', linewidth=1)
    ax1.plot(abf.sweepX, abf.sweepC,'k')#[len(abf.sweepX) * 8:len(abf.sweepX) * 9], 'b', linewidth=1)

    for spine in plt.gca().spines.values():
        spine.set_visible(False)

    plt.gca().get_yaxis().set_visible(False)  # hide Y axis
    plt.gca().get_xaxis().set_visible(False)  # hide X axis
    plt.axis([1.5, 2.85, np.amin(abf.data[1][len(abf.sweepX) * 8:len(abf.sweepX) * 9]),
              np.amax(abf.sweepC)]) # plt.axis([xmin,xmax,ymin,ymax])
    plt.gcf()
    plt.draw()
    fig.savefig(mydirectory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'neuron_IV_profile.png', dpi=1000)
    plt.show()


    print("Successfully completed.")


    # Writting process #####################################################################################################

    Rslt['Filename'] = filename
    Rslt['Resistance Average (MOhm)'] = ResistanceAverage
    Rslt['Resistance Sag (MOhm)'] = SagResistance
    Rslt['Resistance Steady-state (MOhm)'] = SteadyResistance
    Rslt['Baseline Average (mV)'] = np.mean(BaselineAverage)
    Rslt['Baseline Average before Onset (mV)'] = VBaseline[0]
    Rslt['Steady state depolarisation (mV)'] = Vsteady[0]
    Rslt['Sag Max Value (mV)'] = VHcurrent[0]
    Rslt['Sag Amplitude (mV)'] = abs(VHcurrent[0] - Vsteady[0])
    Rslt['Sag Ratio'] = SagRatio
    Rslt['Sag Peak (mV)'] = SagPeak                    # From baseline. Not 'correct' name. Rename if you want.
    Rslt['Sag Full-Width Half maximum (ms)'] = SagHalfWidth
    Rslt['SagFit_a'] = popta[0]
    Rslt['SagFit_b'] = popta[1]
    Rslt['SagFit_c'] = popta[2]


    col = 0
    for thing in Rslt.keys():
        worksheet.write(row+1, col, Rslt[thing])
        col += 1
    row = row + 1

workbook.close()

# Ending ##############################################################################################################
end = time.time()
print("Execution time: ", end - start, 'second(s) - ', (end - start) / len(src_files), 'second(s)/file')  # Beat it!
print("I-V Analysis Done. That's all folks!")
