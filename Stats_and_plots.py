##########################################################################
# Importing modules ######################################################
##########################################################################

import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
from scipy.stats import wilcoxon
import pandas as pd  # Alternatively, you can use xlrd

# Setting global parameters for plot design ##############################
plt.rcParams['axes.linewidth'] = 3  # set the value globally
plt.rcParams["font.family"] = "Calibri"  # Change the font used in plots

##########################################################################
# Import and read your data ##############################################
##########################################################################

df = pd.read_excel(r'C:\Users\River\Desktop\project.xlsx', sheet_name=0, usecols='A,B',  # Path of your xlsx
                   header=None, skiprows=1)  # Read data from xlsx; skip first row with labels

# Set labels of your data
DataLabel = 'Persistent firing frequency (spikes/sec)'  # Or 'Depolarization (mV)' ; Label of the Y-axis

x = df[0]  # your first set of data here
xlabel = 'Carbachol'  # Label of the first column

y = df[1]  # your second set of data here
ylabel = 'ML204'  # Label of the second column

##########################################################################
# Checking the normality of the distributions ############################
##########################################################################

DifferenceData = np.subtract(x, y)
stats.probplot(DifferenceData, dist="norm", plot=plt)
plt.title("Normal Q-Q plot")
plt.savefig(xlabel + '_' + ylabel + '_QQPlot.png', dpi=400)  # Saved where your code is
plt.show()

# Testing the normality of the distributions #############################
s0, p0 = stats.shapiro(DifferenceData)  # Shapiro-Wilk; Check if a better test is available for your data
print('Shapiro-Wilk test:', s0, p0)

# Interpret the results and process test the samples ####################
alpha = 0.05  # Define your confidence interval for Shapiro-Wilk test

if p0 <= alpha:
    print('The null hypothesis (H0) has been rejected. Sample does not seem to be drawn from a normal distribution')
    test = 'Wilcoxon signed-rank test'
    w, p = wilcoxon(x, y)
    print('Wilcoxon signed-rank test:', w, p)
else:
    print('The null hypothesis (H0) failed to be rejected. Sample seems to be drawn from a normal distribution')
    test = "Paired Student's t-test"
    t, p = stats.ttest_rel(x, y)
    print("Paired Student's t-test:", t, p)

##########################################################################
# Plot the values and stats ##############################################
##########################################################################

# plot parameters
w = 0.2  # bar width
colors = ['k', 'k']  # corresponding colors
xcoor = [1, 1.3]  # x-coordinates of your bars
data = [x, y]  # data series

# Plot design
fig, ax = plt.subplots()
ax.bar(xcoor,
       height=[np.mean(yi) for yi in data],
       yerr=[(0, 0), [stats.sem(yi) for yi in data]],  # error bars
       error_kw=dict(lw=3, capsize=15, capthick=3),  # error bar details
       width=w,  # bar width
       tick_label=[xlabel, ylabel],
       color=(0, 0, 0, 0),  # face color transparent
       edgecolor=colors,
       linewidth=4.0,
       )

ax.spines['right'].set_color('none')  # Eliminate upper and right axes
ax.spines['top'].set_color('none')

# distribute scatter over the center of the bars
ax.scatter(np.ones(x.size), x, color='none', edgecolor='grey', s=150, linewidth=2.5)
ax.scatter(np.full((1, y.size), xcoor[1]), y, color='none', edgecolor='grey', s=150, linewidth=2.5)

for element in np.linspace(0, len(data[0]) - 1, len(data[0])):
    ax.plot([xcoor[0], xcoor[1]], [x[int(element)], y[int(element)]], 'k', alpha=0.5)

ax.axhline(y=0, color='k', linestyle='-', linewidth=2)  # when negative values, trace a line at y=0

# Statistical significance levels
if p < 0.001:
    ax.axhline(np.amax(data) + 1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
    ax.text((xcoor[0] + xcoor[1]) / 2, np.amax(data) + 1.2, '***', fontsize=40, horizontalalignment='center')
    sig = str('***')
elif p < 0.01:
    ax.axhline(np.amax(data) + 1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
    ax.text((xcoor[0] + xcoor[1]) / 2, np.amax(data) + 1.2, '**', fontsize=40, horizontalalignment='center')
    sig = str('**')
elif p < 0.05:
    ax.axhline(np.amax(data) + 1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
    ax.text((xcoor[0] + xcoor[1]) / 2, np.amax(data) + 1.2, '*', fontsize=40, horizontalalignment='center')
    sig = str('*')
else:
    ax.axhline(np.amax(data) + 1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
    ax.text((xcoor[0] + xcoor[1]) / 2 - 0, np.amax(data) + 1.2, 'n.s.', fontsize=40, horizontalalignment='center')
    sig = str('n.s.')

# Axes design
plt.ylabel(DataLabel, fontsize=30, fontweight='bold')
plt.xlabel(test + ', ' + sig + r'. $\it{p}$ = %.3f' % p + r', $\it{n}$=' + str(len(data[0])), fontsize=20)
ax.xaxis.set_tick_params(labelsize=30)
ax.xaxis.labelpad = 20
ax.yaxis.set_tick_params(labelsize=30)
ax.yaxis.labelpad = 0

ax.minorticks_on()
ax.yaxis.set_tick_params(which='minor', length=5, width=2, direction='out')

ax.xaxis.set_tick_params(width=3)
ax.yaxis.set_tick_params(width=3)

fig.set_size_inches(10, 13)

fig.savefig(xlabel + '_' + ylabel + '_test_Barplot.png', dpi=400)  # Saved where your code is
plt.show()
