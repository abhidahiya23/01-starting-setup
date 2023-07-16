#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
import xlsxwriter


# In[54]:


# TYpe 2 Pie CHart
# need to calculate the values now i am using mock values from csv
totalhcp = 10435 
df = pd.read_csv('sample2.csv')
approved = 0 
pending = 0
rejected = 0 
nofeedback = 0
for ind in df.index:
    if df['Name'][ind] == 'Approved':
        approved = df['Count'][ind]
    elif  df['Name'][ind] == 'Pending':
        pending = df['Count'][ind]
    elif  df['Name'][ind] == 'Rejected':
        rejected = df['Count'][ind]

nofeedback = totalhcp - approved - pending - rejected
approved_per = round(100*(approved/totalhcp),2)
pending_per = round(100*(pending/totalhcp),2)
rejected_per = round(100*(rejected/totalhcp),2)
nofeeback_per = round(100*(nofeedback/totalhcp),2)

df.loc[len(df.index)] = ["Total HCP",10435]


# In[3]:


print(approved_per)
print(pending_per)
print(rejected_per)
print(nofeeback_per)


# In[50]:


y = np.array([approved_per,pending_per,rejected_per,nofeeback_per])
z = np.array([approved,pending,rejected,nofeedback])
mylabels = ["Approved","Pending","Rejected","NoFeeback"]

#y = plt.pie(y, labels = mylabels)
#y.set_title('hello')
#fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))
#plt.savefig("myImagePDF.pdf", format="pdf", bbox_inches="tight")
#plt.show()


# In[172]:


fig, ax = plt.subplots(figsize=(11.69,8.27), subplot_kw=dict(aspect="equal"))

ingredients = mylabels

def func(pct, allvals):
    absolute = int(np.round(pct/100.*np.sum(allvals)))
    return f"{pct:.2f}%"

wedges, texts, autotexts = ax.pie(z, autopct=lambda pct: func(pct, z),
                                  textprops=dict(color="w"))

ax.legend(wedges, ingredients,
          title="Status",
          loc="upper right")
#          bbox_to_anchor=(0.5,0.5,2,0))

plt.setp(autotexts, size=8, weight="bold")
ax.set_title("Sample 2 Heading")
plt.savefig("myImagePDF.pdf", format="pdf", bbox_inches="tight")
plt.show()

figs=[]
figs.append(fig)


# In[68]:


df2 = pd.read_csv('sample3.csv',delimiter='|')


# In[69]:


df2


# In[173]:


# Generating Stacked Bar Chart 
import matplotlib.pyplot as plt
import numpy as np

areas=[]
weights={    
    "tab 1": np.array([df2['tab 1'][0], df2['tab 1'][1]]),
    "HVBI": np.array([df2['HVBI'][0], df2['HVBI'][1]]),
    "HVF ASM I": np.array([df2['HVF ASM I'][0], df2['HVF ASM I'][1]]),
    "MVI": np.array([df2['MVI'][0], df2['MVI'][1]]),
    "SUPT" : np.array([df2['SUPT'][0], df2['SUPT'][1]])
        }
# data from https://allisonhorst.github.io/palmerpenguins/
for i in range(len(df2.index)):
    #print(df2['Area'][i])
    areas.append(df2['Area'][i])
    #areas
#print(weights)

width = 0.25

fig, ax = plt.subplots(figsize=(11.69,8.27))
#fig.set_figheight(8)
#fig.set_figwidth(8)
bottom = np.zeros(2)

min_value = 10000000000000
max_value = 0

for boolean, weights in weights.items():
    p = ax.bar(areas, weights, width, label=boolean, bottom=bottom)
    bottom += weights
    min_value = min(min_value,bottom[0])
    max_value = max(max_value,bottom[0])

#print(min_value)
#print(max_value)
ax.set_yticks([500, 1000, 1500, 2000, 2500, 3000,3500, 4000, 4500])
ax.yaxis.grid(True, linestyle='-', which='major',
                   color='grey', alpha=.2)
ax.set_title("Segment Distribution at Area Level")
ax.set_ylabel('Segment Count')
ax.set_xlabel('Area')
ax.legend(title="Status",loc="upper right")
plt.savefig("myImagePDF.pdf", format="pdf", bbox_inches="tight")
plt.show()
figs.append(fig)


# In[174]:


from matplotlib.backends.backend_pdf import PdfPages
with PdfPages('multipage_pdf.pdf') as pdf:
    for fig in figs:
        pdf.savefig(fig)
        plt.close()


# In[176]:


# Writing Sheets into Excel Code. Improve it using the below code
"""https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html#:~:text=To%20write%20to%20multiple%20sheets,necessary%20to%20save%20the%20changes."""
df1 = pd.DataFrame(
   [[5, 2], [4, 1]],
   columns=["Rank", "Subjects"]
)

df2 = pd.DataFrame(
   [[15, 21], [41, 11]],
   index=["One", "Two"],
   columns=["Rank", "Subjects"]
)

print(df1)
print(df2)

with pd.ExcelWriter('FeedbackSummary.xlsx') as writer:
    df1.to_excel(writer, sheet_name='Nation Level Territory Feedback',index=False)
    df2.to_excel(writer, sheet_name='Nation Level HCP Feedback',index=False)


# In[ ]:




