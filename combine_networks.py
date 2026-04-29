import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from collections import defaultdict
import math

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

def build_graph_for_department(df_dept, min_weight=2):
    df_dept = df_dept.dropna(subset=['经办人', '代办人', '查验员姓名'])
    edge_weights = defaultdict(int)
    
    for _, row in df_dept.iterrows():
        deputy = row['代办人']
        agent = row['经办人']
        inspector = row['查验员姓名']
        edge_weights[(deputy, agent)] += 1
        edge_weights[(deputy, inspector)] += 1
    
    edge_weights = {k: v for k, v in edge_weights.items() if v > min_weight}
    G = nx.Graph()
    for (u, v), w in edge_weights.items():
        G.add_edge(u, v, weight=w)
    G.remove_nodes_from(list(nx.isolates(G)))
    return G, edge_weights, df_dept

def draw_combined_network():
    df = pd.read_excel('代办情况.xlsx', engine='openpyxl')
    departments = df['业务部门'].unique()
    valid_depts = []
    for dept in departments:
        df_dept = df[df['业务部门'] == dept]
        G, weights, df_dept = build_graph_for_department(df_dept, min_weight=2)
        if G.number_of_nodes() > 0:
            valid_depts.append((dept, G, weights, df_dept))

    n = len(valid_depts)
    if n == 0:
        print("无有效数据")
        return
    cols = min(3, n)
    rows = math.ceil(n / cols)
    fig, axes = plt.subplots(rows, cols, figsize=(7*cols, 7*rows), squeeze=False)
    fig.suptitle('全市各业务部门代办-经办-查验关系网状图 (>2笔)', fontsize=22)

    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor='#1f77b4', label='经办人'),
        Patch(facecolor='#ff7f0e', label='代办人'),
        Patch(facecolor='#2ca02c', label='查验员'),
        Patch(facecolor='#9467bd', label='经办+代办'),
        Patch(facecolor='#8c564b', label='经办+查验'),
        Patch(facecolor='#e377c2', label='代办+查验'),
        Patch(facecolor='#d62728', label='三者皆是'),
        Patch(facecolor='red', label='最密切关系'),
        Patch(facecolor='#444444', label='其他关系 (>2笔)'),
    ]

    for idx, (dept, G, weights, df_dept) in enumerate(valid_depts):
        row = idx // cols
        col = idx % cols
        ax = axes[row, col]

        agents = set(df_dept['经办人'].dropna().unique())
        deputies = set(df_dept['代办人'].dropna().unique())
        inspectors = set(df_dept['查验员姓名'].dropna().unique())
        node_colors = []
        for node in G.nodes():
            is_agent = node in agents
            is_deputy = node in deputies
            is_inspector = node in inspectors
            combo = (is_agent, is_deputy, is_inspector)
            if combo == (True, True, True):
                node_colors.append('#d62728')  # 三者皆是：深红
            elif combo == (True, True, False):
                node_colors.append('#9467bd')  # 经办+代办：紫
            elif combo == (True, False, True):
                node_colors.append('#8c564b')  # 经办+查验：棕
            elif combo == (False, True, True):
                node_colors.append('#e377c2')  # 代办+查验：粉
            elif is_agent:
                node_colors.append('#1f77b4')  # 纯经办人：蓝
            elif is_inspector:
                node_colors.append('#2ca02c')  # 纯查验员：绿
            elif is_deputy:
                node_colors.append('#ff7f0e')  # 纯代办人：橙
            else:
                node_colors.append('#7f7f7f')  # 未知：灰

        max_weight = max(weights.values()) if weights else 0
        edges = list(G.edges())
        edge_widths = []
        edge_colors_list = []
        red_marked = False
        for u, v in edges:
            w = weights.get((u, v), weights.get((v, u), 0))
            width = 1 + 6 * (w / max_weight) if max_weight > 0 else 1
            edge_widths.append(width)
            if w == max_weight and max_weight > 0 and not red_marked:
                edge_colors_list.append('red')
                red_marked = True
            else:
                edge_colors_list.append('#444444')

        pos = nx.spring_layout(G, k=3.0, iterations=150, seed=42)
        nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=250, alpha=0.9, ax=ax)
        nx.draw_networkx_edges(G, pos, width=edge_widths, edge_color=edge_colors_list, alpha=0.5, ax=ax)
        nx.draw_networkx_labels(G, pos, font_size=6, font_family='Microsoft YaHei', ax=ax)
        edge_labels = {(u, v): str(w) for (u, v), w in weights.items()}
        nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=5, font_color='#555555', ax=ax)

        ax.set_title(f'{dept}\n节点:{G.number_of_nodes()} 边:{G.number_of_edges()}', fontsize=13)
        ax.axis('off')

    for idx in range(n, rows*cols):
        row = idx // cols
        col = idx % cols
        axes[row, col].axis('off')

    fig.legend(handles=legend_elements, loc='lower center', ncol=5, fontsize=13)
    plt.tight_layout(rect=[0, 0.05, 1, 0.96])
    plt.savefig('combined_network.png', dpi=300, bbox_inches='tight')
    print(f"已保存为 combined_network.png")

if __name__ == '__main__':
    draw_combined_network()
