import os
import streamlit as st
import pandas as pd
import plotly.express as px




os.environ['LANG'] = 'zh_CN.UTF-8'
def calculate_entry_counts(data):
    # 计算每个上级部门和社会化服务点组合的条目数量
    entry_counts = data.groupby(['上级部门', '社会化服务点']).size().reset_index(name='条目数量')
    return entry_counts

def update_service_points(department):
    # 根据选定的上级部门更新社会化服务点选项
    service_points = st.session_state['entry_counts'][st.session_state['entry_counts']['上级部门'] == department]['社会化服务点'].unique()
    return service_points

def main():
    data = pd.read_excel("代办情况.xlsx", engine='openpyxl')
    entry_counts = calculate_entry_counts(data)
    st.write("Excel数据已加载！")
    # 将数据保存到session state中
    st.session_state['entry_counts'] = entry_counts

    # 获取申请日期的最小值和最大值
    min_apply_date = data['申请日期'].min()
    max_apply_date = data['申请日期'].max()

    # 创建上级部门的选择器
    departments = entry_counts['上级部门'].unique()
    department_selector = st.selectbox("选择上级部门", departments)

    # 根据选定的上级部门更新社会化服务点选项
    service_points = update_service_points(department_selector)
    service_point_selector = st.selectbox("选择社会化服务点", service_points)
    # 创建开始日期和结束日期的选择器，使用申请日期的最小值和最大值作为选项
    start_date_selector = st.date_input("选择开始日期", value=min_apply_date,min_value=min_apply_date, max_value=max_apply_date)
    end_date_selector = st.date_input("选择结束日期", value=max_apply_date, min_value=min_apply_date, max_value=max_apply_date)
    # 将用户输入的日期转换为Pandas Timestamp
    start_date_selector = pd.Timestamp(start_date_selector)
    end_date_selector = pd.Timestamp(end_date_selector)
    # 根据选择筛选数据并进行分类汇总
    filtered_data = data[(data['上级部门'] == department_selector) & (data['社会化服务点'] == service_point_selector) & (data['申请日期'] >= start_date_selector) & (data['申请日期'] <= end_date_selector)]
    grouped_data = filtered_data.groupby(['经办人', '代办人姓名']).size().reset_index(name='次数')
    print(grouped_data)
    # 创建散点图
    
    
    fig = px.scatter(grouped_data, x='代办人姓名', y='经办人', size='次数', color='次数', hover_data=['代办人姓名', '经办人'])



    # 显示散点图
    st.plotly_chart(fig,use_container_width=True)

    with st.container():
        st.markdown("""<footer style="text-align: center; padding: 20px;font-size:12px">
            ©2024张硕|保留所有权利|网站使用<a href="https://www.streamlit.io/" target="_blank">Streamlit</a>构建
        </footer>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()