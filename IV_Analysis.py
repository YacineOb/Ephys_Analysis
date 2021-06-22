""" What this code is doing:
- Calculates Baseline for ALL sweeps and the average Baseline.
- Calculates Sag peaks for ALL sweeps if possible.
- Calculates Sag Amplitude for ALL sweeps if possible.
- Calculates Sag Ratio for ALL sweeps if possible.
- Calculates Sag Exp. fitting + Tau constant for first sweep.
- Calculates the firing frequency for ALL sweeps when Vmb > Vrest.
- Calculates the instantaneous firing frequency and Exp. fit for ALL sweeps when Vmb > Vrest.
- Calculates The Resistance for each sweep and average resistance.
- Calculates the ISI Adaptation """

########################################################################################################################
# Import modules #######################################################################################################
########################################################################################################################
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

########################################################################################################################
# Introduction #########################################################################################################
########################################################################################################################

start = time.time()
timestr = time.strftime("%H-%M-%S.%d.%m.%Y")  # Create file with time

# Path and file manipulations ##########################################################################################

directory = 'C:/Python/InVitro/Analysis_IV'  # Change the working directory. Will be use to find
os.chdir(directory)  # all the files we want to analyze. Change it if you want.

if not os.path.exists(directory + '/Results_IV/' + timestr):  # If no PlotsIv directory exists
    os.makedirs(directory + '/Results_IV/' + timestr)  # We creat it to store our plots in.
    print("Results path has been created")  # display message if created.
else:  # Otherwise, if the directory exists,
    pass  # Just move on.

src_files = [i for i in os.listdir(directory)  # If it finds a directory, will ignore it.
             if not os.path.isdir(i)]

###############################################################################################################
# Definitions #################################################################################################
###############################################################################################################

def monoexp(t, a, b, c):
    return a * np.exp(-t / b) - c


def biexp(t, a, b, c, d, e):
    return a * np.exp(- t / b) - c * np.exp(- t / d) - e

version = 1
if version == 0:
    Cst = {'PulseStart':1.50, 'PulseEnd': 2.50, 'ResistancePulseStart': 4.00,'RPulseInterval': 0.500,
       'ResistancePulseEnd': 4.50, 'SagFrame': 0.150,'AverageOffOnSet': 0.200}
else:
    Cst = {'PulseStart':1.65, 'PulseEnd': 2.65, 'ResistancePulseStart': 4.65,'RPulseInterval': 0.500,
       'ResistancePulseEnd': 5.15, 'SagFrame': 0.165,'AverageOffOnSet': 0.200}

Rslt = {'Filename': 0, 'Resistance Average (MOhm)': 0, 'Resistance Sag (MOhm)': 0, 'Resistance Steady-state (MOhm)': 0,
        'Baseline Average (mV)': 0, 'Baseline Average before Onset (mV)': 0,
        'Steady state depolarisation (mV)': 0, 'Sag Max Value (mV)': 0, 'Sag Amplitude (mV)': 0, 'Sag Ratio': 0,
        'Sag Peak (mV)': 0,'Sag Full-Width Half maximum (ms)': 0, 'SagFast': 0, 'SagSlow': 0, 'SagMono': 0,
        'Rebound Depo (mV)':0, 'Rebound FWHM (ms)': 0}

''' Variables in dict.
Var = {'Vsteady': 0,'VHcurrent':0, 'ResistanceSweep':0, 'ResistanceSweep':0, 'BaselineSweep':0,
            'VBaseline':0, 'PeaksSweep':0, 'Frequency':0, 'InstantaneousFrequency':0}
'''
########################################################################################################################
# Excel file opening ###################################################################################################
########################################################################################################################

workbook = xlsxwriter.Workbook(directory + '/Results_IV/' + timestr + '/Results_' + timestr + '.xlsx')
worksheet = workbook.add_worksheet('Basic_properties')

col = 0
row = 0
for key in Rslt.keys():
    worksheet.write(row, col, key)  # row, col, item
    col += 1

########################################################################################################################
# ABF file opening #####################################################################################################
########################################################################################################################

for filename in src_files:
    os.makedirs(directory + '/Results_IV/' + timestr + '/' + filename[:-4])
    abf: ABF = pyabf.ABF(filename)

    print("Reading", filename + '...', end=""),

    ####################################################################################################################
    # ABF-dependant Variables definitions ##############################################################################
    ####################################################################################################################

    # Convert constants to sample points
    Cst = {key: int(Cst[key] * abf.dataRate) for key in Cst.keys()}

    # ! I create the current input data. Use SweepC for ABF2.0. No SweepC with WinWCP converted ABF
    CurrentIn = np.linspace(-300, 300, len(abf.sweepList))


    # Define variables
    Vsteady = np.zeros(len(abf.sweepList))
    VHcurrent = ResistanceSweep = np.zeros(len(abf.sweepList))
    ResistanceSweep = np.zeros(len(abf.sweepList))
    BaselineSweep = np.zeros(len(abf.sweepList))
    VBaseline = np.zeros(len(abf.sweepList))
    Frequency = np.zeros(len(abf.sweepList))
    PeaksSweep = np.zeros(len(abf.sweepList), dtype=object)
    ReboundDepo = np.zeros(len(abf.sweepList), dtype=object)
    ReboundDepo_half = np.zeros(len(abf.sweepList), dtype=object)
    InstantaneousFrequency = np.zeros(len(abf.sweepList), dtype=object)
    AdaptationRatio = np.zeros(len(abf.sweepList), dtype=object)

    ####################################################################################################################
    # Algorithm core - Main iteration ##################################################################################
    ####################################################################################################################

    for sweepNumber in abf.sweepList:
        abf.setSweep(sweepNumber)

        # Baseline from 0 second to the beginning of the pulse.
        BaselineSweep[sweepNumber] = np.mean(abf.sweepY[:Cst['PulseStart']])  # Mean/baseline before current pulse
        VBaseline[sweepNumber] = np.mean(abf.sweepY[Cst['PulseStart'] - Cst['AverageOffOnSet']:Cst['PulseStart']])
        ResistanceSweep[sweepNumber] = (np.mean(abf.sweepY[Cst['ResistancePulseStart']:Cst['ResistancePulseEnd']])
                                        - np.mean(
                    abf.sweepY[Cst['ResistancePulseStart'] - Cst['RPulseInterval']:Cst['ResistancePulseStart']])) \
                                       / (np.mean(
            abf.sweepC[Cst['ResistancePulseStart']:Cst['ResistancePulseEnd']]) * 1e-3)  # real values in data var.
        # deltaV for Resistance

        # If the mean during the pulse frame is bellow the baseline +5 mV, we calculate the sag.
        if np.mean(abf.sweepY[Cst['PulseStart']:Cst['PulseEnd']]) <= BaselineSweep[sweepNumber] :
            VHcurrent[sweepNumber] = np.amin(
                abf.sweepY[Cst['PulseStart']:Cst['PulseStart'] + Cst['SagFrame']])  # Take the minimum Voltage for each sweep.
            Vsteady[sweepNumber] = np.mean(abf.sweepY[
                                           Cst['PulseEnd'] - Cst['AverageOffOnSet']:Cst['PulseEnd']])


            ReboundDepo[sweepNumber], _ = find_peaks(abf.sweepY,prominence=2,width = 500,distance=100000)
            ReboundDepo_half[sweepNumber] = peak_widths(abf.sweepY, ReboundDepo[sweepNumber], rel_height=0.5)

            plt.plot(abf.sweepX, abf.sweepY)
            plt.plot(abf.sweepX[ReboundDepo[sweepNumber][0]],abf.sweepY[ReboundDepo[sweepNumber][0]],'xr')
            plt.show()


        # If the mean during the pulse frame is equal to the baseline, there is probably no sag and no spikes.
        elif np.mean(abf.sweepY[Cst['PulseStart']:Cst['PulseEnd']]) == BaselineSweep[sweepNumber] + 5:
            continue

        # If the mean during the pulse frame is above the baseline, count spikes.
        else:
            PeaksSweep[sweepNumber], _ = find_peaks(abf.sweepY, height=-5)
            Frequency[sweepNumber] = len(PeaksSweep[sweepNumber])

    print(" Processing data...", end="")

    ####################################################################################################################
    # Basic properties layout ##########################################################################################
    ####################################################################################################################

    if math.isnan(float(np.mean(ResistanceSweep))) or math.isinf(np.mean(ResistanceSweep)) == True:
        warnings.warn(filename + " doesn't have -5pA input. Resistance is set to 0 ")
        ResistanceAverage = 0
    else:
        ResistanceAverage = np.mean(ResistanceSweep)

    BaselineAverage = np.mean(BaselineSweep)
    SagPeak = BaselineSweep[0] - VHcurrent[0]  # From baseline.
    SagRatio = (VHcurrent[0] - Vsteady[0]) / (VHcurrent[0] - VBaseline[0])


    ISIthreshold = list(map(lambda i: i >= 8, Frequency)).index(True)  # Take all sweeps ∋ nSpike > 8. For ISI adapt.
    for index, item in enumerate(PeaksSweep[ISIthreshold:]):
        AdaptationRatio[index + ISIthreshold] = [(item[-1] - item[-2]) / (item[1] - item[0])]

    fthreshold = list(map(lambda i: i > 3, Frequency)).index(True)  # Take all sweeps ∋ nSpike > 3. For fitting.
    for index, item in enumerate(PeaksSweep[fthreshold:]):
        InstantaneousFrequency[index + fthreshold] = [1 / (item[j + 1] - item[j]) * abf.dataRate for j in
                                                      np.arange(0, len(item) - 1)]
    '''
        # Fitting Instantaneous frequency process
        p0 = [-100, 1, 100, 10, 10]  # The function need some help on the parameters. input your first guess here.
        popt, pcov = optimize.curve_fit(biexp, np.arange(0, len(PeaksSweep[index + fthreshold]) - 1),
                                        InstantaneousFrequency[index + fthreshold], p0, maxfev=5000)
        # print("a =", popt[0], "+/-", pcov[0, 0] ** 0.5)
        # print("b =", popt[1], "+/-", pcov[1, 1] ** 0.5)
        # print("c =", popt[2], "+/-", pcov[2, 2] ** 0.5)
        yEXP = biexp(np.arange(0, len(PeaksSweep[index + fthreshold]) - 1), *popt)

        # Plotting processed data ######################################################################################

        # Instantaneous frequency and fitting plot
        fig = plt.figure()
        ax = fig.add_subplot(111)
        ax.spines['left'].set_linewidth(2)
        ax.spines['right'].set_color('none')
        ax.spines['top'].set_color('none')
        ax.spines['bottom'].set_color('none')
        plt.gca().get_xaxis().set_visible(False)  # hide X axis
        plt.scatter(range(0, len(PeaksSweep[index + fthreshold]) - 1), InstantaneousFrequency[index + fthreshold],
                    label='Data', facecolors='None', edgecolors='k')
        plt.plot(range(0, len(PeaksSweep[index + fthreshold]) - 1), yEXP, 'r-', ls='--',
                 label='a=%5.1f, b=%5.1f, c=%5.1f, d=%5.1f, e=%5.1f' % tuple(popt))
        plt.text(0.3, 0.75, r'$F_{instantaneous}(t)=a{e}^{-b{t}}+c{e}^{-d{t}}-e$', fontsize=15,
                 transform=plt.gca().transAxes)  # Disable/Delete if you don't use Tex.
        plt.title("Instantaneous frequency - " + str(filename) + "; \n $n_{sweep} =$ " + str(index + fthreshold))
        plt.ylabel('Instantaneous frequency (Hz)', fontweight='bold')
        plt.legend()
        fig.savefig(directory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/'
                    + 'Instantaneous_frequency_sweep_' + str(index + fthreshold) + '.png', dpi=400)
        plt.show()
    '''

    ####################################################################################################################
    # Plotting processed data ##########################################################################################
    ####################################################################################################################

    # Plotting the first sag trace, baseline, sag amplitude, ratio and Vsteady versus Vsag #############################
    abf.setSweep(abf.sweepList[0])  # We want to plot the first trace only

    ####################################################################################################################
    # I-V Plot #########################################################################################################
    ####################################################################################################################

    fig = plt.figure()
    ax = fig.add_subplot(111)

    plt.scatter(CurrentIn[:len(np.trim_zeros(Vsteady))], np.trim_zeros(Vsteady),
                label='Steady state', facecolors='white', edgecolors='k')
    plt.plot(CurrentIn[:len(np.trim_zeros(Vsteady))], scipy.signal.savgol_filter(np.trim_zeros(Vsteady), 5, 2),
             linewidth=1.5, linestyle='--', color='k', label='Steady state fitted')  # , marker='o')

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
    # plt.axhline(y=0, linewidth=1, color='k', linestyle='dotted')
    plt.title("Current-Potential relationship")
    plt.ylabel('Potential (mV)', fontweight='bold')
    plt.xlabel('Current Injected (pA)', fontweight='bold')
    plt.legend()
    fig.savefig(
        directory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Current-Potential_relationship.png',
        dpi=400)

    plt.show()

    # Plotting the frequency of sweeps #################################################################################

    fig = plt.figure()
    ax = fig.add_subplot(111)
    for axis in ['bottom', 'left']:
        ax.spines[axis].set_linewidth(2)

    plt.scatter(CurrentIn, Frequency,
                label='Steady state', facecolors='w', edgecolors='k')
    plt.plot(CurrentIn, scipy.signal.savgol_filter(Frequency, 5, 3), linewidth=1.5,
             linestyle='--', color='k')  # , marker='o')

    ax.spines['right'].set_color('none')  # Eliminate upper and right axes
    ax.spines['top'].set_color('none')
    plt.xlim(left=-50)  # adjust the left leaving right unchanged
    plt.title("Current-Frequency relationship")
    plt.xlabel('Current Injected (pA)', fontweight='bold')
    plt.ylabel('Average Firing Frequency (Hz)', fontweight='bold')
    fig.savefig(directory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Frequency.png', dpi=400)
    # plt.show()

    # FWHM######################################################################################
    #PUT THIS ABOVE?
    y = abf.sweepY[Cst['PulseStart']:Cst['PulseEnd']]
    x = abf.sweepX[Cst['PulseStart']:Cst['PulseEnd']]
    c = abf.sweepC[Cst['PulseStart']:Cst['PulseEnd']]

    peaksPC, _ = find_peaks(-y, prominence=(5), distance=5000)
    results_half = peak_widths(-y, peaksPC, rel_height=0.5)

    SagHalfWidth = results_half[0] / 20000 * 10e2   # keep all values?
    # print("half W:", SagHalfWidth,
    #      "peak:", peaksPC)

    # plt.plot(x*20000, y, 'k', label='Data')
    # plt.plot(1.65*20000+peaksPC, y[peaksPC], "rx")
    # plt.hlines(results_half[1]*-1,1.65*20000+results_half[2],1.65*20000+results_half[3], color='red')
    # plt.xlim([Cst['PulseStart'],Cst['PulseEnd']])
    # plt.plot(x[int(peaksPC)-100:int(peaksPC)+100]*20000,y[int(peaksPC)-100:int(peaksPC)+100])#,color='blue')
    # plt.show()

    # Resistance (MOhm)#######################################

    SagPotential = np.mean(y[int(peaksPC) - 100:int(peaksPC) + 100])
    InputCurrent = np.mean(c[int(peaksPC) - 100:int(peaksPC) + 100])
    SagResistance = (VBaseline[0] - (SagPotential)) / -(InputCurrent * 1e-3)
    SteadyResistance = (VBaseline[0] - Vsteady[0]) / \
                       -(np.mean(abf.sweepC[Cst['PulseEnd'] - Cst['AverageOffOnSet']:Cst['PulseEnd']]) * 1e-3)

    # Plot sag with ratio########################

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
                 xytext=(2.70, Vsteady[0] + (Vsteady[0] - VHcurrent[0]) * - 0.45))  # Sag Ratio
    plt.annotate(str(round(Vsteady[0] - VHcurrent[0], 2)) + ' mV', xy=(2.70, Vsteady[0]),
                 xytext=(2.70, VHcurrent[0] + (VHcurrent[0] - Vsteady[0]) * -0.25))  # Vsag - Vsteady

    plt.annotate('', xy=(1.50, VHcurrent[0]), xytext=(1.50, BaselineSweep[0]), arrowprops=dict(arrowstyle='<->'))
    plt.annotate(str(round(SagPeak, 2)) + ' mV', xy=(1.28, VHcurrent[0]),
                 xytext=(1.45, (VHcurrent[0] + BaselineSweep[0]) / 2), rotation=90, va='center')  # Sag Amplitude
    plt.annotate(str(round(np.mean(BaselineSweep), 2)) + ' mV', xy=(1.25, BaselineSweep[0]),
                 xytext=(1.50, BaselineSweep[0] + 1.25), ha='center')  # Baseline

    plt.axhline(y=VHcurrent[0], linewidth=0.7, color='r', linestyle='--')
    plt.axhline(y=Vsteady[0], linewidth=0.7, color='b', linestyle='--')
    plt.axhline(y=BaselineSweep[0], linewidth=0.7, color='k', linestyle='--')

    plt.hlines(results_half[1] * -1, (1.65 * 20000 + results_half[2]) / 20000, (1.65 * 20000 + results_half[3]) / 20000,
               color='red')

    plt.title("Sag properties")
    plt.xlabel('Time (s)', fontweight='bold')
    plt.ylabel('Potential (mV)', fontweight='bold')
    fig.savefig(directory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'Sag_properties.png',
                dpi=400)  # Disable if you don't need to plot. Increases speed.
    plt.show()

    ####################################################################################################################
    # Sag magnification and exponential fitting ########################################################################
    ####################################################################################################################

    # Define the area where the sag should be ##########################################################################
    IndexSag = np.amin(np.where(abf.sweepY == VHcurrent[0]))

    y1 = abf.sweepY[int(1.69 * abf.dataRate):int(2.65 * abf.dataRate)]
    x1 = np.linspace(0, 2.65 - 1.69, len(y1))
    #x1 = abf.sweepX[int(IndexSag+0.007):int(2.65 * abf.dataRate)]
    #y1 = abf.sweepY[int(IndexSag-0.007):int(2.65 * abf.dataRate)-1]

    # Biexponential fitting ############################################################################################
    p1 = [VHcurrent[0], 0.1, VHcurrent[0], 0.5, Vsteady[0]]
    popt1, pcov1 = optimize.curve_fit(biexp, x1, y1, p1, maxfev=10000)
    yEXP1 = biexp(x1, *popt1)

    # Monoexponential fitting ##########################################################################################
    p2 = [-80, 0.1, 5]  # Change if you think the dynamic is different.
    popt2, pcov2 = optimize.curve_fit(monoexp, x1, y1, p2, maxfev=10000)  # Big number of iteration.
    yEXP2 = monoexp(x1, *popt2)

    # Fitting plotting #################################################################################################
    fig = plt.figure()  # Prepare the figure
    ax = fig.add_subplot(111)
    plt.plot(x1, y1, 'grey', label='Data')
    plt.xlabel('Time (s)', fontweight='bold')
    plt.ylabel('Potential (mV)', fontweight='bold')
    for axis in ['bottom', 'left']:
        ax.spines[axis].set_linewidth(2)
    ax.spines['right'].set_color('none')
    ax.spines['top'].set_color('none')

    # Plot Bi-exponential #############################################
    plt.plot(x1, yEXP1, 'r-', ls='--', label='a=%5.3f, Tau_1=%5.3f, c=%5.3f, Tau_2=%5.3f, e=%5.3f' % tuple(popt1))
    plt.text(0.55, 0.35, r'$V(t)=a{e}^{-b{t}}+c{e}^{-d{t}}-e$', color='red', fontsize=10,
             transform=plt.gca().transAxes)

    # Plot Mono-exponential #############################################
    plt.plot(x1, yEXP2, 'b-', ls='--', label='a=%5.3f, Tau_0=%5.3f, c=%5.3f' % tuple(popt2))
    plt.text(0.55, 0.3, r'$V(t)=a{e}^{-b{t}}-c$', color='blue', fontsize=10,
             transform=plt.gca().transAxes)

    # Display the legend (parameters), save and show plot ############
    plt.legend()
    fig.savefig(directory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'SagFit.png', dpi=400)
    plt.show()

    ####################################################################################################################
    # Plotting a figure with the first and last sweep.##################################################################
    ####################################################################################################################

    # Define the TOP layout ############################################################################################
    gs = gridspec.GridSpec(2, 1, height_ratios=[10, 1])
    axs = plt.subplots(2, 1, sharex=True,gridspec_kw={'wspace': 0, 'hspace': 0})  # Remove horizontal space between axes
    ax0 = plt.subplot(gs[0])  # Plot each graph, and manually set the y tick values
    for spine in plt.gca().spines.values():
        spine.set_visible(False)
    ax0.annotate(str(round(np.mean(BaselineSweep), 2)) + ' mV', xy=(1.2, np.mean(BaselineSweep)),
                 xytext=(1.2, np.mean(BaselineSweep) + 10), fontsize=10, fontweight='bold', ha='center')# change baselinesweep value with a simple variable
    plt.gca().get_yaxis().set_visible(False)  # hide Y axis
    plt.gca().get_xaxis().set_visible(False)  # hide X axis

    # Plot blue trace on top
    abf.setSweep(abf.sweepList[-9])  # Take the last sweep BUT CHANGE THIS!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    limitmax = np.amax(abf.sweepY)   # This trace define the max
    ax0.plot(abf.sweepX, abf.sweepY, 'blue', linewidth=1)

    #Plot grey trace on top
    abf.setSweep(abf.sweepList[0])   # Take the first sweep
    limitmin = np.amin(abf.sweepY) # This trace define the min
    ax0.plot(abf.sweepX, abf.sweepY, 'gray', linewidth=1)

    #Add legend (voltage)
    ax0.plot([2.70, 2.80], [limitmax-abs(limitmax*1.2), limitmax-abs(limitmax*1.2)], 'k', linewidth=2, zorder=10)
    ax0.plot([2.70, 2.70], [limitmax-abs(limitmax*1.2), limitmax-abs(limitmax*1.2) + 15], 'k', linewidth=2, zorder=10)
    ax0.annotate('15 mV', xy=(2.72, limitmax-abs(limitmax*0.8)), xytext=(2.72, limitmax-abs(limitmax*0.8)),
                 fontsize=7, fontweight='bold',zorder=4)
    ax0.annotate('100 ms', xy=(2.72, limitmax-abs(limitmax*1.1)), xytext=(2.72, limitmax-abs(limitmax*1.1)),
                 fontsize=7, fontweight='bold',zorder=4)

    # Restrain the view to the pulse step
    plt.axis([1.5, 2.85, limitmin, limitmax])  # plt.axis([xmin,xmax,ymin,ymax])


    # Define the BOTTOM layout #########################################################################################
    ax1 = plt.subplot(gs[1])
    for spine in plt.gca().spines.values():
        spine.set_visible(False)
    plt.gca().get_yaxis().set_visible(False)  # hide Y axis
    plt.gca().get_xaxis().set_visible(False)  # hide X axis

    # Plot current inputs
    ax1.plot(abf.sweepX, abf.data[1][len(abf.sweepX) * 15:len(abf.sweepX) * 16], 'b', linewidth=2)
    ax1.plot(abf.sweepX, abf.sweepC, 'k')

    # Add legend (current)
    ax1.plot([2.70, 2.70], [100,400], 'k', linewidth=2, zorder=10)
    ax1.annotate('300 pA', xy=(2.72, 150), xytext=(2.72, 150),
                 fontsize=7, fontweight='bold',zorder=4)

    # Restrain the view to the pulse step
    plt.axis([1.5, 2.85, np.amin(abf.sweepC),
              np.amax(abf.data[1][len(abf.sweepX) * 15:len(abf.sweepX) * 16])]) # plt.axis([xmin,xmax,ymin,ymax])

    # Save and show
    plt.savefig(directory + '/Results_IV/' + timestr + '/' + filename[:-4] + '/' + 'neuron_IV_profile.png', dpi=1000)
    plt.show()

    ####################################################################################################################
    ####################################################################################################################
    ####################################################################################################################

    print("Successfully completed.")

    ####################################################################################################################
    # Excel writing process ############################################################################################
    ####################################################################################################################

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
    Rslt['Sag Peak (mV)'] = SagPeak  # From baseline.
    Rslt['Sag Full-Width Half maximum (ms)'] = SagHalfWidth
    Rslt['SagFast'] = popt1[1]
    Rslt['SagSlow'] = popt1[3]
    Rslt['SagMono'] = popt2[1]
    Rslt['Rebound Depo (mV)'] = ReboundDepo[0][0]  # From baseline. Not 'correct' name. Rename if you want.
    Rslt['Rebound FWHM (ms)'] = ReboundDepo_half[0][0]/20000*1000
    # Pass to next row for next file ###################################################################################
    col = 0
    for thing in Rslt.keys():
        worksheet.write(row + 1, col, Rslt[thing])
        col += 1
    row = row + 1

workbook.close()

########################################################################################################################
# Ending ###############################################################################################################
########################################################################################################################

end = time.time()
print("Execution time: ", end - start, 'second(s) - ', (end - start) / len(src_files), 'second(s)/file')
print("I-V Analysis Done. That's all folks!")
