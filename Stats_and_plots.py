

# Importing modules #########################################
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
import numpy as np
from scipy import stats
from scipy.stats import wilcoxon


# Setting global parameters #################################
plt.rcParams['axes.linewidth'] = 3 # set the value globally
plt.rcParams["font.family"] = "Calibri" # Change the font used in plots


# Import and read your data ############################################

#depoCch = np.array([1.026504517,1.334060669,0.243732452,1.332427979])
#depoML204 = np.array([-1.61208725,0.036701202, 0.899219513,0.190662384])

#depoCch = np.array([1.026504517,1.334060669,0.243732452,1.332427979,0.642166138,5.482227325,1.961933136,5.896385193])
#depoML204 = np.array([-1.61208725,0.036701202, 0.899219513,0.190662384,-1.102390289,1.386692047,1.290462494,4.274814606])

depoCch = np.array([9,1.6,2.5,3.3,1.7,12.9,5.8,6.7])
depoML204 = np.array([2,0.1,2.1,3.2,0,4.2,0,5.9])



k2, p0 = stats.normaltest(depoCch)
print("p0 = {:g}".format(p0))

shapiro_test = stats.shapiro(depoCch)
print(shapiro_test)

w, p = wilcoxon(depoCch,depoML204)
print(w,p)

t,p1 = stats.ttest_rel(depoCch,depoML204)
print(t,"This is p1", p1)

w = 0.2    # bar width
x = [1, 1.3] # x-coordinates of your bars
colors = ['k', 'k']    # corresponding colors
y = [depoCch,       # data series
    depoML204]

figure(num=None, figsize=(8, 9), dpi=400, facecolor='w', edgecolor='k')
fig, ax = plt.subplots()
ax.bar(x,
       height=[np.mean(yi) for yi in y],
       yerr=[stats.sem(yi) for yi in y],    # error bars
       capsize=10, # error bar cap width in points
       width=w,    # bar width
       tick_label=["Carbachol", "ML204"],
       color=(0,0,0,0),  # face color transparent
       edgecolor=colors,
       linewidth = 4.0,
       #ecolor=colors,    # error bar colors; setting this raises an error for whatever reason.
       )
ax.spines['right'].set_color('none')  # Eliminate upper and right axes
ax.spines['top'].set_color('none')


    # distribute scatter randomly across whole width of bar
ax.scatter(np.ones(depoCch.size),depoCch, color= 'none',edgecolor = 'grey', s=150,linewidth = 2.5)
ax.scatter(np.full((1,depoML204.size),x[1]),depoML204, color= 'none',edgecolor = 'grey', s=150,linewidth = 2.5)


for element in np.linspace(0,len(y[0])-1,len(y[0])):
       ax.plot([x[0],x[1]] , [depoCch[int(element)], depoML204[int(element)]], 'k', alpha = 0.5)


ax.axhline(y=0, color='k', linestyle='-', linewidth = 2)

if p < 0.001:
       ax.axhline(np.amax(y)+1,0.25,0.75, color='k', linestyle='-',linewidth = 2)
       ax.text((x[0]+x[1])/2, np.amax(y)+1.2,'***',fontsize=40,horizontalalignment='center')
       sig = str('***')
elif p < 0.01:
       ax.axhline(np.amax(y)+1,0.25,0.75, color='k', linestyle='-',linewidth = 2)
       ax.text((x[0]+x[1])/2, np.amax(y)+1.2,'**',fontsize=40,horizontalalignment='center')
       sig = str('**')
elif p < 0.05:
       ax.axhline(np.amax(y)+1,0.25,0.75, color='k', linestyle='-',linewidth = 2)
       ax.text((x[0]+x[1])/2, np.amax(y)+1.2,'*',fontsize=40,horizontalalignment='center')
       sig = str('*')
else:
       ax.axhline(np.amax(y) + 1, 0.25, 0.75, color='k', linestyle='-',linewidth = 2)
       ax.text((x[0] + x[1]) / 2-0, np.amax(y) + 1.2, 'n.s.', fontsize=40,horizontalalignment='center')
       sig = str('n.s.')


plt.ylabel('Depolarization (mV)', fontsize=30, fontweight='bold')
plt.xlabel('Wilcoxon test, ' + sig + '. p=' + str(p) + ', n=' + str(len(y[0])), fontsize=17)
ax.xaxis.set_tick_params(labelsize=25)
ax.xaxis.labelpad = 20
ax.yaxis.set_tick_params(labelsize=25)
ax.yaxis.labelpad = 0

ax.minorticks_on()
ax.yaxis.set_tick_params(which='minor', length=5, width=2, direction='out')

ax.xaxis.set_tick_params(width=3)
ax.yaxis.set_tick_params(width=3)

fig.set_size_inches(7, 10)
#fig.savefig('test2png.png', dpi=100)

plt.show()