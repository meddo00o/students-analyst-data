import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

file=pd.read_csv('student_exam_scores.csv')
file=file.sort_values(ascending=False,by='exam_score').reset_index(drop=True)
def made_sheet(col,name):
    name=file.groupby('student_id')[str(col)].sum().reset_index(name=str(col)).sort_values(ascending=False,by=str(col)).reset_index(drop=True)
    return name



def add_grade(x):
    if 45<=x<60:
        grade='A'
    elif 40<=x<45:
        grade='B'
    elif 35<=x<40:
        grade='C'
    elif 30<=x<35:
        grade="D"
    elif x<30:
        grade='F'
    return grade

def add_relation(page,x_axis,y_axis,bg):
    page = px.scatter(
        file,
        x=x_axis,
        y=y_axis,
        trendline='ols'
    )
    #control in properties
    page.update_layout(
        title={
            'text': f'Relationship between {x_axis} and {y_axis}<br><sup>Report: There is a positive correlation between {x_axis} and {y_axis}.</sup>',
            'x': 0.5,  # توسيط العنوان
            'xanchor': 'center'
        }
    ,paper_bgcolor=bg)
    page.data[1].line.update(color='red', width=4)
    page.write_html(f'Relationship between {x_axis} and {y_axis}.html',config={'displaylogo':False})

add_relation('fig','hours_studied','exam_score',"#CC61AE")
add_relation('fig','sleep_hours','exam_score',"#61BECC")
add_relation('fig','attendance_percent','exam_score',"#D4A665")
add_relation('fig','previous_scores','exam_score',"#65D4AD")


file['grade']=file['exam_score'].astype(float)
file['grade']=file['grade'].apply(add_grade).sort_values(ascending=False)
print(file)


with pd.ExcelWriter('result.xlsx',engine='openpyxl')as writer:
    file.to_excel(writer,sheet_name='all',index=False)
    made_sheet(col='hours_studied',name='df1').to_excel(writer,sheet_name='hours_studied',index=False)
    made_sheet(col='sleep_hours',name='df2').to_excel(writer,sheet_name='sleep_hours',index=False)
    made_sheet(col='attendance_percent',name='df3').to_excel(writer,sheet_name='attendance_percent',index=False)
    made_sheet(col='previous_scores',name='df4').to_excel(writer,sheet_name='previous_scores',index=False)
    made_sheet(col='exam_score',name='df5').to_excel(writer,sheet_name='exam_score',index=False)


