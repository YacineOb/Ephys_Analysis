# Importing modules #########################################################
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
import numpy as np
from scipy import stats
from scipy.stats import wilcoxon

# Setting global parameters for plot design #################################
plt.rcParams['axes.linewidth'] = 3 # set the value globally
plt.rcParams["font.family"] = "Calibri" # Change the font used in plots

# Import and read your data #################################################

######################################################################################################################
#Read data from excel?
PFcch = 1
PFML204 = 1
###################################################################

DataLabel = 'Persistent firing frequency (spikes/sec)'  # Or 'Depolarization (mV)' ; Label of the Y-axis

x = PFCch
xlabel = 'Carbachol'   # Label of the first column

y = PFML204
ylabel = 'ML204'       # Label of the second column

# Checking the normality of the distributions #################################
DifferenceData = np.subtract(x, y)
stats.probplot(DifferenceData, dist="norm", plot=plt)
plt.title("Normal Q-Q plot")
plt.savefig(xlabel + '_' + ylabel + '_QQPlot.png', dpi=1000)  # Saved where your code is
plt.show()

# Testing the normality of the distributions ####################################
s0, p0 = stats.shapiro(DifferenceData)  # Shapiro-Wilk; Robust but check if a better test is available for your data
print('Shapiro-Wilk test:', s0, p0)

# Interpret the results and process test the samples #######################################
alpha = 0.05     # Define your confidence interval for Shapiro-Wilk test

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


# Plot the values and stats ##############################################

w = 0.2    # bar width
colors = ['k', 'k']    # corresponding colors

xcoor = [1, 1.3]  # x-coordinates of your bars
data = [x, y]  # data series

figure(num=None, figsize=(8, 9), dpi=400, facecolor='w', edgecolor='k')
fig, ax = plt.subplots()
ax.bar(xcoor,
       height=[np.mean(yi) for yi in data],
       yerr=[(0,0),[stats.sem(yi) for yi in data]],    # error bars
       error_kw=dict(lw=3, capsize=15, capthick=3),    # error bar details
       width=w,    # bar width
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


for element in np.linspace(0, len(data[0])-1, len(data[0])):
       ax.plot([xcoor[0], xcoor[1]], [x[int(element)], y[int(element)]], 'k', alpha=0.5)


ax.axhline(y=0, color='k', linestyle='-', linewidth=2)

if p < 0.001:
       ax.axhline(np.amax(data)+1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
       ax.text((xcoor[0]+xcoor[1])/2, np.amax(data)+1.2, '***', fontsize=40, horizontalalignment='center')
       sig = str('***')
elif p < 0.01:
       ax.axhline(np.amax(data)+1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
       ax.text((xcoor[0]+xcoor[1])/2, np.amax(data)+1.2, '**', fontsize=40, horizontalalignment='center')
       sig = str('**')
elif p < 0.05:
       ax.axhline(np.amax(data)+1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
       ax.text((xcoor[0]+xcoor[1])/2, np.amax(data)+1.2, '*', fontsize=40, horizontalalignment='center')
       sig = str('*')
else:
       ax.axhline(np.amax(data) + 1, 0.25, 0.75, color='k', linestyle='-', linewidth=2)
       ax.text((xcoor[0] + xcoor[1]) / 2-0, np.amax(data) + 1.2, 'n.s.', fontsize=40, horizontalalignment='center')
       sig = str('n.s.')


plt.ylabel(DataLabel, fontsize=30, fontweight='bold')
plt.xlabel(test + ', ' + sig + '. p = %.3f' % p + ', $\it{n}$=' + str(len(data[0])), fontsize=17)
#plt.xlabel("Paired Student's t-test", ' + sig + '. p= %.3f' % p + ', n=' + str(len(y[0])), fontsize=17)
ax.xaxis.set_tick_params(labelsize=25)
ax.xaxis.labelpad = 20
ax.yaxis.set_tick_params(labelsize=25)
ax.yaxis.labelpad = 0

ax.minorticks_on()
ax.yaxis.set_tick_params(which='minor', length=5, width=2, direction='out')

ax.xaxis.set_tick_params(width=3)
ax.yaxis.set_tick_params(width=3)

fig.set_size_inches(7, 10)
fig.savefig(xlabel + '_' + ylabel + '_test_Barplot.png', dpi=1000)  # Saved where your code is
plt.show()
