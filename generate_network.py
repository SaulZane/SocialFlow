import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from collections import defaultdict

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

def build_graph_for_department(df_dept):
    df_dept = df_dept.dropna(subset=['经办人', '代办人', '查验员姓名'])
    
    edge_weights = defaultdict(int)
    for _, row in df_dept.iterrows():
        agent = row['经办人']
        deputy = row['代办人']
        inspector = row['查验员姓名']
        if pd.notna(agent) and pd.notna(deputy):
            edge_weights[(agent, deputy)] += 1
        if pd.notna(agent) and pd.notna(inspector):
            edge_weights[(agent, inspector)] += 1
        if pd.notna(deputy) and pd.notna(inspector):
            edge_weights[(deputy, inspector)] += 1
    
    # 过滤 ≤5 笔的关系
    edge_weights = {k: v for k, v in edge_weights.items() if v > 5}
    
    # 过滤后移除孤立节点
    G = nx.Graph()
    for (u, v), w in edge_weights.items():
        G.add_edge(u, v, weight=w)
    
    # 移除没有任何边的孤立节点
    G.remove_nodes_from(list(nx.isolates(G)))
    
    return G, edge_weights, df_dept

def draw_department_network(G, edge_weights, df_dept, dept_name, output_path):
    if G.number_of_nodes() == 0:
        print(f"  [{dept_name}] 无可显示数据（所有关系均≤5笔），跳过")
        return

    agents = set(df_dept['经办人'].dropna().unique())
    deputies = set(df_dept['代办人'].dropna().unique())
    inspectors = set(df_dept['查验员姓名'].dropna().unique())

    node_colors = []
    for node in G.nodes():
        if node in agents:
            node_colors.append('#1f77b4')
        elif node in inspectors:
            node_colors.append('#2ca02c')
        elif node in deputies:
            node_colors.append('#ff7f0e')
        else:
            node_colors.append('#7f7f7f')

    max_weight = max(edge_weights.values()) if edge_weights else 0

    edge_widths = []
    edge_colors_list = []
    for (u, v), w in edge_weights.items():
        width = 1 + 8 * (w / max_weight) if max_weight > 0 else 1
        edge_widths.append(width)
        if w == max_weight and max_weight > 0:
            edge_colors_list.append('red')
        else:
            edge_colors_list.append('#444444')

    pos = nx.spring_layout(G, k=3.0, iterations=150, seed=42)

    fig, ax = plt.subplots(figsize=(24, 20))
    nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=300, alpha=0.9, ax=ax)
    nx.draw_networkx_edges(G, pos, width=edge_widths, edge_color=edge_colors_list, alpha=0.5, ax=ax)
    nx.draw_networkx_labels(G, pos, font_size=6, font_family='Microsoft YaHei', ax=ax)

    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor='#1f77b4', label=f'经办人 ({len([n for n in G.nodes() if n in agents])})'),
        Patch(facecolor='#ff7f0e', label=f'代办人 ({len([n for n in G.nodes() if n in deputies])})'),
        Patch(facecolor='#2ca02c', label=f'查验员 ({len([n for n in G.nodes() if n in inspectors])})'),
        Patch(facecolor='red', label=f'最密切关系 (权重={max_weight})'),
        Patch(facecolor='#444444', label='其他关系 (>5笔)'),
    ]
    ax.legend(handles=legend_elements, loc='upper left', fontsize=12)
    ax.set_title(f'{dept_name} 代办-经办-查验关系网状图 (>5笔)\n节点数: {G.number_of_nodes()}, 边数: {G.number_of_edges()}', fontsize=18)
    ax.axis('off')
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"  [{dept_name}] 已保存 -> {output_path}")

def main():
    df = pd.read_excel('代办情况.xlsx', engine='openpyxl')
    departments = df['业务部门'].unique()
    
    print(f"共 {len(departments)} 个部门，开始分别生成网络图...")
    for dept in departments:
        df_dept = df[df['业务部门'] == dept]
        G, weights, df_dept = build_graph_for_department(df_dept)
        safe_name = dept.replace('/', '_')
        output_path = f'network_{safe_name}.png'
        draw_department_network(G, weights, df_dept, dept, output_path)

if __name__ == '__main__':
    main()
