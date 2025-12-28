import networkx as nx
import random
import numpy as np
import pandas as pd
from openpyxl import Workbook
from matplotlib import pyplot as plt


# 级联模型所需参数

# 项目数量
project_number = 60
# 项目投资额度
project_investment = np.array([[102, 62, 63, 66, 109, 120, 103, 70, 113, 74, 117, 66, 96, 73, 107, 86, 113, 67, 68, 78,
                                81, 115, 118, 105, 105, 92, 100, 75, 94, 92, 106, 116, 90, 105, 66, 80, 105, 91, 73, 76,
                                120, 119, 115, 101, 90, 93, 84, 120, 91, 92, 110, 75, 64, 94, 100, 73, 117, 93, 118, 62]])
# 企业能力目标
capacity_number = 20
# 企业能力权重
capacity_weight = np.array([[0.08, 0.03, 0.01, 0.08, 0.07, 0.07, 0.06, 0.06, 0.04, 0.06, 0.05, 0.04, 0.02, 0.03, 0.05,
                             0.06, 0.03, 0.03, 0.03, 0.07]])
# 项目节点权重
node_weight = np.zeros((1, project_number))
node_weight_1 = np.zeros((1, project_number))  # 考虑资源时候使用
# 项目节点的最初承载力计算
initial_node_load = np.zeros((1, project_number))
initial_node_load_1 = np.zeros((1, project_number))  # 考虑资源时候使用
# 项目节点的最初容量计算
initial_node_capacity = np.zeros((1, project_number))
initial_node_capacity_1 = np.zeros((1, project_number))  # 考虑资源时候使用
# 项目成功概率
project_success_rate = np.zeros((1, project_number))


# 基础参数设定
# 级联失效层和是否中断控制参数

# 初始负载调整参数1-alph
adfactor_1 = 0.1
# 初始负载调整参数2-beta
adfactor_2 = 0.8
# 初始负载调整参数1-eta1
adfactor_1_1 = 0.5
# 初始负载调整参数2-eta2
adfactor_2_2 = 0.5
# 容量调整参数1-lamda
adcapacity_1 = 0.5
# 容量调整参数2-gama
adcapacity_2 = 1.1
# 项目对能力的支撑程度矩阵
file_path = 'P_S_C.xlsx'
P_S_C = np.array(pd.read_excel(file_path, engine='openpyxl'))
# 项目对能力是否存在支撑矩阵，属于判断矩阵
file_path = 'P_S_C_0_1.xlsx'
P_S_C_0_1 = np.array(pd.read_excel(file_path, engine='openpyxl'))
# 能力间关系
file_path = 'Relationship_C.xlsx'  # 替换为你的Excel文件路径
Relationship_C = np.array(pd.read_excel(file_path, engine='openpyxl'))  # 能力与能力间关系，行对列有帮助

# 项目的“输入-处理-输出”过程
file_path = 'paper.xlsx'
df = np.array(pd.read_excel(file_path, engine = 'openpyxl'))
# 项目信息存储列表，每个元素为字典
P = []

for i in range(df.shape[0]):
    di = {}
    di["input"] = list(df[i][1: 6])
    di["process"] = list(df[i][6: 9])
    di["output_new"] = list(df[i][9: 12])
    di["output_recycle"] = list(df[i][12: 14])
    di["start_time"] = list(df[i][14: 15])
    di["end_time"] = list(df[i][15: 16])
    di["strategy"] = list(df[i][16: 36])
    P.append(di)
# 去除Excel中关于0的信息
for i in range(df.shape[0]):
    dd = P[i]
    while 0 in dd["input"]:
        dd["input"].remove(0)
    while 0 in dd["output_new"]:
        dd["output_new"].remove(0)
    while 0 in dd["process"]:
        dd["process"].remove(0)
    while 0 in dd["output_recycle"]:
        dd["output_recycle"].remove(0)
    while 0 in dd["start_time"]:
        dd["start_time"].remove(0)
    while 0 in dd["end_time"]:
        dd["end_time"].remove(0)
    while 0 in dd["strategy"]:
        dd["strategy"].remove(0)

# 依赖关系确定函数

# 技术依赖关系
def Project_interdependency_knowledge(P):
    Rule = ["input", "process", "output_new", "output_recycle", "time", "strategy"]
    project_num = len(P)
    PIs_K_matrix = np.zeros((project_num, project_num))
    i_times = 0
    for i in P:
        j_times = 0
        for j in P:
            if i ==j:
                continue
            else:
                if i['end_time'] <= j['start_time']:
                    D1 = len(set(i[Rule[1]]) & set(j[Rule[0]]))
                    D2 = len(set(i[Rule[1]]) & set(j[Rule[1]]))
                    if (D1 + D2) > 0:
                        PIs_K_matrix[i_times][j_times] = 1
                    else:
                        PIs_K_matrix[i_times][j_times] = 0
                j_times = j_times + 1
        i_times = i_times + 1
    return PIs_K_matrix


# 资源依赖关系
def Project_interdependency_resource(P):
    Rule = ["input", "process", "output_new", "output_recycle", "time", "strategy"]
    project_num = len(P)
    PIs_R_matrix = np.zeros((project_num, project_num))
    i_times = 0
    for i in P:
        j_times = 0
        for j in P:
            if i ==j:
                continue
            else:
                if i['end_time'] <= j['start_time']:
                    D1 = len(set(i[Rule[3]]) & set(j[Rule[0]]))
                    if D1 > 0:
                        PIs_R_matrix[i_times][j_times] = 1
                if i['start_time'] == j['start_time']:
                    D2 = len(set(i[Rule[0]]) & set(j[Rule[0]]))
                    if D2 > 0:
                        PIs_R_matrix[i_times][j_times] = 1
                j_times = j_times + 1
        i_times = i_times + 1
    return PIs_R_matrix


# 战略依赖关系
def Project_interdependency_strategy(P):
    Rule = ["input", "process", "output_new", "output_recycle", "time", "strategy"]
    project_num = len(P)
    PIs_S_matrix = np.zeros((project_num, project_num))
    i_times = 0
    for i in P:
        j_times = 0
        for j in P:
            if i==j:
                continue
            else:
                D1 = len(set(i[Rule[5]]) & set(j[Rule[5]]))
                if D1 > 0:
                    PIs_S_matrix[i_times][j_times] = 1
                else:
                    PIs_S_matrix[i_times][j_times] = 0
                j_times = j_times + 1
        i_times = i_times + 1
    return PIs_S_matrix


# 成果依赖关系
def Project_interdependency_outcome(P):
    Rule = ["input", "process", "output_new", "output_recycle", "time", "strategy"]
    project_num = len(P)
    PIs_OUT_matrix = np.zeros((project_num, project_num))
    i_times = 0
    for i in P:
        j_times = 0
        for j in P:
            if i==j:
                continue
            else:
                if i['end_time'] <= j['start_time']:
                    D1 = len(set(i[Rule[2]]) & set(j[Rule[0]]))
                    D2 = len(set(i[Rule[2]]) & set(j[Rule[1]]))
                    if (D1 + D2) > 0:
                        PIs_OUT_matrix[i_times][j_times] = 1
                    else:
                        PIs_OUT_matrix[i_times][j_times] = 0
                j_times = j_times + 1
        i_times = i_times + 1
    return PIs_OUT_matrix


# 集合相似度计算函数
def Jaccard(list1, list2):
    intersection = len(list(set(list1).intersection(list2)))
    union = (len(list1) + len(list2)) - intersection
    return float(intersection) / union

np.set_printoptions(precision = 5)

# 依赖关系测度
P_PIs_matrix_1 = Project_interdependency_outcome(P)
P_PIs_matrix_2 = Project_interdependency_resource(P)
P_PIs_matrix_3 = Project_interdependency_knowledge(P)
P_PIs_matrix_4 = Project_interdependency_strategy(P)

# 项目间依赖关系强度矩阵
P_P_matrix_1 = P_PIs_matrix_4 + P_PIs_matrix_3 + P_PIs_matrix_2 + P_PIs_matrix_1
# 由于项目间依赖关系和影响关系的方向是相反的，需要对项目间依赖矩阵进行转置
P_P_matrix = P_P_matrix_1.T

# 参数控制区
interrupt_project_rate = 1.0
zhongduan = 1
IP = 0
cc_rate = []
cc_rate_average = []
rou = []
iteration = 3000
# 最初项目失败节点
num_initial_failures = 7

# 资源分配策略1
Y_1 = 0.33
# 资源分配策略2
Y_2 = 0.33
# 资源分配策略3
Y_3 = 0.33
# 终断成本系数
eta = 0

while IP < iteration:
    # 迭代次数统计
    IP = IP + 1
    # 项目层网络-G
    G = nx.from_numpy_array(P_P_matrix, create_using=nx.DiGraph)
    G_star = nx.from_numpy_array(P_P_matrix, create_using=nx.DiGraph)
    # 图的介数
    GJN = nx.betweenness_centrality(G_star)

    # 项目层级联失效过程
    # 初始负载确定
    # 节点权重确定
    # 确定节点度
    degree_node = nx.degree(G_star)

    sum_DG = 0
    for node in list(G_star.nodes()):
        sum_DG = sum_DG + np.sum(P_P_matrix, 0)[node] + np.sum(P_P_matrix, 1)[node]

    for node in list(G.nodes()):
        nei_node_for_weight = G.neighbors(node)
        w_node_i = adfactor_1_1 * ((np.sum(P_P_matrix, 0)[node] + np.sum(P_P_matrix, 1)[node]) / sum_DG) + adfactor_2_2 * (project_investment[0][node] / np.sum(project_investment))  # 对应公式（3）
        node_weight[0][node] = w_node_i

    for i in range(project_number):
        node_weight_1[0][i] = node_weight[0][i]


    # 确定节点初始负载和容量
    for node in list(G.nodes()):
        nei_node = list(G.neighbors(node))
        load_node_i = (1 + adfactor_1) * (node_weight[0][node]) ** adfactor_2  # 节点i的初始负载
        initial_node_load[0][node] = float(load_node_i)
        node_capacity = initial_node_load[0][node] + adcapacity_1 * (initial_node_load[0][node] ** adcapacity_2)
        initial_node_capacity[0][node] = node_capacity

    for i in range(project_number):
        initial_node_capacity_1[0][i] = initial_node_capacity[0][i]


    # 考虑中断的项目级联失效模型
    if zhongduan == 1:
        # 可理解为项目执行比例
        # 项目剩余资源的分配
        # 初始崩溃节点
        initial_failed_nodes = []
        jjj = 0
        while jjj < num_initial_failures:
            jjj = jjj + 1
            cc = random.randint(0, 59)
            while cc in initial_failed_nodes:
                cc = random.randint(0, 59)
            initial_failed_nodes.append(cc)

        # 更新节点容量所需要的节点权重
        resource_matrix = np.zeros((1, project_number))
        resource_matrix_j = np.zeros((1, project_number))
        for failed_nodes in list(initial_failed_nodes):
            if initial_node_load[0][failed_nodes] != 0:
                nei_failure_nodes = list(G.neighbors(failed_nodes))
                kechuandi_nodes = list(set(nei_failure_nodes) - set(initial_failed_nodes))
                li_3 = []
                li_4 = []
                li_5 = []
                if kechuandi_nodes != []:
                    for node_j in kechuandi_nodes:
                        total_sum_load_j = initial_node_load[0][node_j]
                        li_3.append(total_sum_load_j)
                        total_impact_other_j = P_P_matrix[failed_nodes][node_j]
                        li_4.append(total_impact_other_j)
                        total_nei_rem_capacity_j = initial_node_capacity[0][node_j] - initial_node_load[0][node_j]
                        li_5.append(total_nei_rem_capacity_j)
                    total_sum_load = sum(li_3)
                    total_impact_other = sum(li_4)
                    total_nei_rem_capacity = sum(li_5)
                    # 项目终断后节点权重
                    if total_nei_rem_capacity > 0: # 因为total_nei_rem_capacity是分母，所以需要对其是否为0进行判定 print("-------------------------------------出错了！！！！--------------------------------------------")
                        for k_node in kechuandi_nodes:
                            resource_matrix[0][k_node] = resource_matrix[0][k_node] + ((1 - interrupt_project_rate - eta) * project_investment[0][failed_nodes] *
                                                        (Y_1 * (initial_node_load[0][k_node] / total_sum_load) +
                                                         Y_2 * (P_P_matrix[failed_nodes][k_node] / total_impact_other) +
                                                         Y_3 * (1 - ((initial_node_capacity[0][k_node] - initial_node_load[0][k_node]) / total_nei_rem_capacity))))
                    else:
                        for k_node in kechuandi_nodes:
                            resource_matrix[0][k_node] = resource_matrix[0][k_node] + ((1 - interrupt_project_rate - eta) * project_investment[0][failed_nodes] *
                                                        (Y_1 * (initial_node_load[0][k_node] / total_sum_load) +
                                                         Y_2 * (P_P_matrix[failed_nodes][k_node] / total_impact_other)+
                                                         Y_3 * len(kechuandi_nodes)))

                    for k_node in kechuandi_nodes:
                        node_weight_1[0][k_node] = adfactor_1_1 * ((np.sum(P_P_matrix, 0)[k_node] + np.sum(P_P_matrix, 1)[k_node]) / sum_DG) + \
                                                   adfactor_2_2 * ((project_investment[0][k_node] + resource_matrix[0][k_node]) / np.sum(project_investment))
                else:
                    continue
            else:
                continue

        # 在T时刻，更新容量
        for node in list(G_star.nodes()):
            nei_node = list(G_star.neighbors(node))
            load_node_i = (1 + adfactor_1) * (node_weight_1[0][node]) ** adfactor_2  # 节点i的初始负载
            initial_node_load_1[0][node] = float(load_node_i)
            node_capacity = initial_node_load_1[0][node] + adcapacity_1 * (initial_node_load_1[0][node] ** adcapacity_2)
            initial_node_capacity[0][node] = node_capacity


        # 更新容量结束后，级联失效模型
        weight_total_matrix = np.zeros((project_number, 1))
        weight_node_matrix = np.zeros((project_number, project_number))
        for failed_nodes in list(initial_failed_nodes):
            if initial_node_load[0][failed_nodes] != 0:
                nei_failure_nodes = list(G.neighbors(failed_nodes))
                kechuandi_nodes = list(set(nei_failure_nodes) - set(initial_failed_nodes))
                li_1 = []
                if kechuandi_nodes != []:
                    for j in kechuandi_nodes:
                        weight_node_matrix[failed_nodes][j] = initial_node_capacity[0][j] - initial_node_load[0][j]
                        if weight_node_matrix[failed_nodes][j] < 0:
                            print("--------------------------------------------有出错的----------------------------------------------")
                        else:
                            li_1.append(weight_node_matrix[failed_nodes][j])
                    weight_total_matrix[failed_nodes][0] = sum(li_1)
                else:
                    continue
            else:
                continue

        for failed_nodes in list(initial_failed_nodes):
            if initial_node_load[0][failed_nodes] != 0:
                nei_failure_nodes = list(G.neighbors(failed_nodes))
                kechuandi_nodes = list(set(nei_failure_nodes) - set(initial_failed_nodes))
                if kechuandi_nodes != []:
                    if weight_total_matrix[failed_nodes][0] > 0:
                        for j in kechuandi_nodes:
                            node_load_update = initial_node_load[0][j] + initial_node_load[0][failed_nodes] * \
                                               (weight_node_matrix[failed_nodes][j] / weight_total_matrix[failed_nodes][0])
                            initial_node_load[0][j] = node_load_update
                    else:
                        for j in kechuandi_nodes:
                            node_load_update = initial_node_load[0][j] + initial_node_load[0][failed_nodes] / len(kechuandi_nodes)
                            initial_node_load[0][j] = node_load_update
                else:
                    continue
            else:
                continue

        G.remove_nodes_from(initial_failed_nodes)

        # 检查是否有新的节点因为负载超过容量而崩溃
        while True:
            new_failed_nodes = set()
            for node in list(G.nodes()):
                if initial_node_load[0][node] > initial_node_capacity[0][node]:
                    new_failed_nodes.add(node)
            # 如果没有新的节点崩溃，则结束模拟
            if not new_failed_nodes:
                break
            else:
                # 如果有，继续实施级联失效模型
                # 增加的部分，使得每一个传递都考虑到了终断，并且该部分的作用是更新每个节点的权重和容量

                # 容量更新
                for failed_nodes in list(new_failed_nodes):
                    if initial_node_load[0][failed_nodes] != 0:
                        nei_failure_nodes = list(G.neighbors(failed_nodes))
                        kechuandi_nodes = list(set(nei_failure_nodes) - set(new_failed_nodes))
                        li_3 = []
                        li_4 = []
                        li_5 = []
                        if kechuandi_nodes != []:
                            for node_j in kechuandi_nodes:
                                total_sum_load_j = initial_node_load[0][node_j]
                                li_3.append(total_sum_load_j)
                                total_impact_other_j = P_P_matrix[failed_nodes][node_j]
                                li_4.append(total_impact_other_j)
                                total_nei_rem_capacity_j = initial_node_capacity[0][node_j] - initial_node_load[0][node_j]
                                li_5.append(total_nei_rem_capacity_j)
                            total_sum_load = sum(li_3)
                            total_impact_other = sum(li_4)
                            total_nei_rem_capacity = sum(li_5)
                            # 节点权重更新为了更新容量 print("-------------------------------------出错了！！！！--------------------------------------------")
                            if total_nei_rem_capacity > 0:
                                for k_node in kechuandi_nodes:
                                    resource_matrix[0][k_node] = resource_matrix[0][k_node] + ((1 - interrupt_project_rate - eta) * project_investment[0][failed_nodes] * \
                                                       (Y_1 * (initial_node_load[0][k_node] / total_sum_load) +
                                                        Y_2 * (P_P_matrix[failed_nodes][k_node] / total_impact_other) +
                                                        Y_3 * (1 - ((initial_node_capacity[0][k_node] - initial_node_load[0][k_node]) / total_nei_rem_capacity))))

                            else:
                                 for k_node in kechuandi_nodes:
                                    resource_matrix[0][k_node] = resource_matrix[0][k_node] + ((1 - interrupt_project_rate - eta) *  project_investment[0][failed_nodes] * \
                                                       (Y_1 * (initial_node_load[0][k_node] / total_sum_load) +
                                                        Y_2 * (P_P_matrix[failed_nodes][k_node] / total_impact_other) +
                                                        Y_3 / len(kechuandi_nodes)))

                            for k_node in kechuandi_nodes:
                                node_weight_1[0][k_node] = adfactor_1_1 * (((np.sum(P_P_matrix, 0)[k_node] + np.sum(P_P_matrix, 1)[k_node])) / sum_DG) + \
                                                           adfactor_2_2 * ((project_investment[0][k_node] + resource_matrix[0][k_node]) / np.sum(project_investment))
                        else:
                            continue
                    else:
                        continue

                # 在T时刻，更新容量
                for node in list(G_star.nodes()):
                    nei_node = list(G_star.neighbors(node))
                    load_node_i = (1 + adfactor_1) * (node_weight_1[0][node]) ** adfactor_2  # 节点i的初始负载
                    initial_node_load_1[0][node] = float(load_node_i)
                    node_capacity = initial_node_load_1[0][node] + adcapacity_1 * (initial_node_load_1[0][node] ** adcapacity_2)
                    initial_node_capacity[0][node] = node_capacity

                # 增加部分完毕
                weight_total_matrix = np.zeros((project_number, 1))
                weight_node_matrix = np.zeros((project_number, project_number))
                for node in list(new_failed_nodes):
                    if initial_node_load[0][node] != 0:
                        nei_new_failed_nodes = G.neighbors(node)
                        kechuandi_nodes_for_new = list(set(nei_new_failed_nodes) - set(new_failed_nodes))
                        li_2 = []
                        if kechuandi_nodes_for_new != []:
                            for j in kechuandi_nodes_for_new:
                                weight_node_matrix[node][j] = initial_node_capacity[0][j] - initial_node_load[0][j]
                                if weight_node_matrix[node][j] < 0: # 为了防止计算出现错误
                                    print("小于0")
                                    print(j)
                                    print(initial_node_capacity[0][j])
                                    print(initial_node_load[0][j])
                                else:
                                    li_2.append(weight_node_matrix[node][j])
                            weight_total_matrix[node][0] = sum(li_2)
                        else:
                            continue
                    else:
                        continue

                for node in list(new_failed_nodes):
                    if initial_node_load[0][node] != 0:
                        nei_new_failed_nodes = G.neighbors(node)
                        kechuandi_nodes_for_new = list(set(nei_new_failed_nodes) - set(new_failed_nodes))
                        if kechuandi_nodes_for_new != []:
                            if weight_total_matrix[node][0] > 0:
                                for j in kechuandi_nodes_for_new:
                                    node_load_update_new = initial_node_load[0][j] + initial_node_load[0][node] * \
                                                           (weight_node_matrix[node][j] / weight_total_matrix[node][0])  # 不选择中断，承载分给其他人
                                    initial_node_load[0][j] = node_load_update_new
                            else:
                                for j in kechuandi_nodes_for_new:
                                    node_load_update_new = initial_node_load[0][j] + initial_node_load[0][node] / len(kechuandi_nodes_for_new)  # 不选择中断，承载分给其他人
                                    initial_node_load[0][j] = node_load_update_new
                        else:
                            continue
                    else:
                        continue

                G.remove_nodes_from(new_failed_nodes)
        if zhongduan == 1:
            print("-----------------------------------终断情景下项目层网络-----------------------------------")
            print(G)
        else:
            print("-----------------------------------忽略终断情景下项目层网络-----------------------------------")
            print(G)

    # 合作者层级联失效到影响战略实现的全过程
    elif zhongduan == 2:
        
        # 合作层网络构建-G2
        file_path = 'cooperator_network.xlsx'  # 替换为你的Excel文件路径
        df = pd.read_excel(file_path, engine='openpyxl')
        df = np.array(df)
        C_num = df.shape[0]
        C_C_matrix = np.zeros((C_num, C_num))

        dic = {}
        for i in range(C_num):
            dic[str(df[i][0])] = list(df[i][1:7])
            while 0 in dic[str(df[i][0])]:
                dic[str(df[i][0])].remove(0)
        for co1 in dic.keys():
            for co2 in dic.keys():
                if Jaccard(dic[co1], dic[co2]) > 0:
                    C_C_matrix[int(co1)][int(co2)] = 1
                if int(co1) == int(co2):
                    C_C_matrix[int(co1)][int(co2)] = 0

        adj_matrix = C_C_matrix

        # 合作者层矩阵保存为Excel
        wb = Workbook()
        ws = wb.active
        for row in C_C_matrix:
            ff = list(row)
            ws.append(ff)
        wb.save("C_C_matrix111.xlsx")

        G_2 = nx.from_numpy_array(adj_matrix, create_using = nx.DiGraph)
        G_2_star = nx.from_numpy_array(adj_matrix, create_using=nx.DiGraph)
        degree_node_E = nx.degree(G_2_star)
        sum_DG_E = 0

        for node in list(G_2.nodes()):
            sum_DG_E = sum_DG_E + (np.sum(adj_matrix, 0)[node] + np.sum(adj_matrix, 1)[node])

        # 公司与各个合作者间的关系

        Relationship_E = np.array([[0.79, 0.88, 0.72, 0.87, 0.85, 0.77, 0.76, 0.62, 0.84, 0.79, 0.75, 0.74, 0.67, 0.95, 0.94,
              0.85, 0.86, 0.84, 0.70, 0.88, 0.75, 0.70, 0.85, 0.81, 0.75, 0.80, 0.75, 0.68, 0.67, 0.68,
              0.80, 0.72, 0.88, 0.77, 0.78, 0.88, 0.76, 0.78, 0.77, 0.65, 0.66]])


        # 合作对项目支撑的判断矩阵
        E_S_P_0_1 = np.zeros((C_num, project_number))

        for i in range(C_num):
            for j in range(7):
                ind = int(df[i][j + 1])
                if ind != 0:
                    E_S_P_0_1[i][ind - 1] = 1


        # 合作者数量
        Num_E = E_S_P_0_1.shape[0]
        # 合作者节点最初负载
        initial_node_load_E = np.zeros((1, Num_E))
        # 合作者层最初容量
        initial_node_capacity_E = np.zeros((1, Num_E))
        # 合作者层节点权重
        node_weight_E = np.zeros((1, Num_E))


    # 级联失效过程
        # 初始负载确定
        for node in list(G_2.nodes()):
            w_node_j = adfactor_1_1 * ((np.sum(adj_matrix, 0)[node] + np.sum(adj_matrix, 1)[node])/sum_DG_E) +\
                       adfactor_2_2 * (1 - Relationship_E[0][node]) / (Num_E - np.sum(Relationship_E))
            node_weight_E[0][node] = w_node_j

        for node in list(G_2.nodes()):
            nei_node = list(G_2.neighbors(node))
            li = []
            for i in nei_node:
                nei_node_weight_i = node_weight_E[0][i]
                li.append(nei_node_weight_i)
            nei_node_weight_E = sum(li)
            load_node_j = (1 + adfactor_1) * (node_weight_E[0][node]) ** adfactor_2  # 节点j的初始负载
            initial_node_load_E[0][node] = float(load_node_j)
            node_capacity_j = initial_node_load_E[0][node] + adcapacity_1 * (initial_node_load_E[0][node] ** adcapacity_2)
            initial_node_capacity_E[0][node] = node_capacity_j

        # 失效仿真模拟
        # 随机生成失败节点
        initial_failed_nodes_E = []
        jjj = 0
        while jjj < num_initial_failures:
            jjj = jjj + 1
            cc = random.randint(0, 40)
            while cc in initial_failed_nodes_E:
                cc = random.randint(0, 40)
            initial_failed_nodes_E.append(cc)

        # 节点负载传递过程
        weight_total_matrix_E = np.zeros((project_number, 1))
        weight_node_matrix_E = np.zeros((project_number, project_number))
        for failed_nodes in list(initial_failed_nodes_E):
            if initial_node_load_E[0][failed_nodes] != 0:
                nei_failure_nodes_E = list(G_2.neighbors(failed_nodes))
                kechuandi_nodes_E = list(set(nei_failure_nodes_E) - set(initial_failed_nodes_E))
                li_1 = []
                if kechuandi_nodes_E != []:
                    for j in kechuandi_nodes_E:
                        weight_node_matrix_E[failed_nodes][j] = initial_node_capacity_E[0][j] - initial_node_load_E[0][j]
                        if weight_node_matrix_E[failed_nodes][j] < 0:
                            print("--------------------------------------出错了1---------------------------------------------")
                        else:
                            li_1.append(weight_node_matrix_E[failed_nodes][j])
                    weight_total_matrix_E[failed_nodes][0] = sum(li_1)
                else:
                    continue
            else:
                continue

        for failed_nodes in list(initial_failed_nodes_E):
            if initial_node_load_E[0][failed_nodes] != 0:
                nei_failure_nodes_E = list(G_2.neighbors(failed_nodes))
                kechuandi_nodes_E = list(set(nei_failure_nodes_E) - set(initial_failed_nodes_E))
                if kechuandi_nodes_E != []:
                    if weight_total_matrix_E[failed_nodes][0] > 0:
                        for j in kechuandi_nodes_E:
                            node_load_update_E = initial_node_load_E[0][j] + initial_node_load_E[0][failed_nodes] * \
                                                 (weight_node_matrix_E[failed_nodes][j] / weight_total_matrix_E[failed_nodes][0])
                            initial_node_load_E[0][j] = node_load_update_E
                    else:
                        for j in kechuandi_nodes_E:
                            node_load_update_E = initial_node_load_E[0][j] + initial_node_load_E[0][failed_nodes] / len(kechuandi_nodes_E)
                            initial_node_load_E[0][j] = node_load_update_E
                else:
                    continue
            else:
                continue

        G_2.remove_nodes_from(initial_failed_nodes_E) # 去除已崩溃节点

        # 检查是否有新的节点因为负载超过容量而崩溃
        while True:
            new_failed_nodes_E = set()
            for node in list(G_2.nodes()):
                if initial_node_load_E[0][node] > initial_node_capacity_E[0][node]:
                    new_failed_nodes_E.add(node)
         # 如果没有新的节点崩溃，则结束模拟
            if not new_failed_nodes_E:
                break
            else:
            # 如果有新崩溃的节点
                weight_total_matrix_E_j = np.zeros((project_number, 1))
                weight_node_matrix_E_j = np.zeros((project_number, project_number))
                for node in list(new_failed_nodes_E):
                    if initial_node_load_E[0][node] != 0:
                        nei_new_failed_nodes_E = G_2.neighbors(node)
                        kechuandi_nodes_for_new_E = list(set(nei_new_failed_nodes_E) - set(new_failed_nodes_E))
                        li_2 = []
                        if kechuandi_nodes_for_new_E != []:
                            for j in kechuandi_nodes_for_new_E:
                                weight_node_matrix_E_j[node][j] = initial_node_capacity_E[0][j] - initial_node_load_E[0][j]
                                if weight_node_matrix_E_j[node][j] < 0:
                                    print("--------------------------------------出错了3-----------------------------------------")
                                else:
                                    li_2.append(weight_node_matrix_E_j[node][j])
                            weight_total_matrix_E_j[node][0] = sum(li_2)
                        else:
                            continue
                    else:
                        continue

                for node in list(new_failed_nodes_E):
                    if initial_node_load_E[0][node] != 0:
                        nei_new_failed_nodes_E = G_2.neighbors(node)
                        kechuandi_nodes_for_new_E = list(set(nei_new_failed_nodes_E) - set(new_failed_nodes_E))
                        if kechuandi_nodes_for_new_E != []:
                            if weight_total_matrix_E_j[node][0] > 0:
                                for j in kechuandi_nodes_for_new_E:
                                    node_load_update_new_E = initial_node_load_E[0][j] + initial_node_load_E[0][node] * \
                                                             (weight_node_matrix_E_j[node][j] / weight_total_matrix_E_j[node][0])
                                    initial_node_load_E[0][j] = node_load_update_new_E
                            else:
                                for j in kechuandi_nodes_for_new_E:
                                    node_load_update_new_E = initial_node_load_E[0][j] + initial_node_load_E[0][node] / len(kechuandi_nodes_for_new_E)
                                    initial_node_load_E[0][j] = node_load_update_new_E
                        else:
                            continue
                    else:
                        continue
                G_2.remove_nodes_from(new_failed_nodes_E)
        print("-----------------------------------------合作层网络-----------------------------------------")
        print(G_2)

        # 合作者层失效对项目层的影响
        # 合作者层对项目的支撑矩阵而不是判断矩阵
        '''# 这是生成算法具体以文件《E_S_P.xlsx》为准
        E_S_P = np.zeros((C_num, project_number))
        for i in range(E_S_P_0_1.shape[0]):
            for j in range(E_S_P_0_1.shape[1]):
                if E_S_P_0_1[i][j] == 1:
                    E_S_P[i][j] = random.uniform(0.5, 1.5)

        wb = Workbook()
        ws = wb.active
        for row in E_S_P:
            ff = list(row)
            ws.append(ff)
        wb.save("E_S_P.xlsx")
        '''
        file_path = 'E_S_P.xlsx'  # 替换为你的Excel文件路径
        E_S_P = np.array(pd.read_excel(file_path, engine='openpyxl'))

        E_S_P_GM = np.zeros((1, project_number)) # 合作者层级联失效后，各个项目所受到的支撑
        E_S_P_star = np.ones((1, project_number)) # 各个项目的阈值关于所需要的合作者的支持程度

        for i in range(project_number):
            E_S_P_star[0][i] = 1

        CM_failed_P_from_E = [] # 因合作者失效所导致的项目节点失效
        for i in list(G.nodes):
            li_E_P = []
            for nodes in list(G_2.nodes):
                li_E_P.append(E_S_P[nodes][i])
            E_S_P_GM[0][i] = sum(li_E_P)
            if E_S_P_GM[0][i] < E_S_P_star[0][i]:
                CM_failed_P_from_E.append(i)

        # 判断出因为合作者失效而失效的项目后，进行项目层的级联失效，为便于计算，所以不考虑项目中断
        weight_total_matrix_P_P = np.zeros((project_number, 1))
        weight_node_matrix_P_P = np.zeros((project_number, project_number))
        for failed_nodes in list(CM_failed_P_from_E):
            if initial_node_load[0][failed_nodes] != 0:
                nei_failure_nodes = list(G.neighbors(failed_nodes))
                kechuandi_nodes = list(set(nei_failure_nodes) - set(CM_failed_P_from_E))
                li_1 = []
                if kechuandi_nodes != []:
                    for j in kechuandi_nodes:
                        weight_node_matrix_P_P[failed_nodes][j] = initial_node_capacity[0][j] - initial_node_load[0][j]
                        if weight_node_matrix_P_P[failed_nodes][j] < 0:
                            print("--------------------------------------出错了5-------------------------------")
                        else:
                            li_1.append(weight_node_matrix_P_P[failed_nodes][j])
                    weight_total_matrix_P_P[failed_nodes][0] = sum(li_1)
                else:
                    continue
            else:
                continue

        for failed_nodes in list(CM_failed_P_from_E):
            if initial_node_load[0][failed_nodes] != 0:
                nei_failure_nodes = list(G.neighbors(failed_nodes))
                kechuandi_nodes = list(set(nei_failure_nodes) - set(CM_failed_P_from_E))
                if kechuandi_nodes != []:
                    if weight_total_matrix_P_P[failed_nodes][0] > 0:
                        for j in kechuandi_nodes:
                            node_load_update = initial_node_load[0][j] + initial_node_load[0][failed_nodes] * \
                                               (weight_node_matrix_P_P[failed_nodes][j] / weight_total_matrix_P_P[failed_nodes][0])  # 不选择中断，承载分给其他人
                            initial_node_load[0][j] = node_load_update
                    else:
                        for j in kechuandi_nodes:
                            node_load_update = initial_node_load[0][j] + initial_node_load[0][failed_nodes] / len(kechuandi_nodes)  # 不选择中断，承载分给其他人
                            initial_node_load[0][j] = node_load_update
                else:
                    continue
            else:
                continue
        G.remove_nodes_from(CM_failed_P_from_E)

        while True:   # 检查是否有新的节点因为负载超过容量而崩溃
            new_failed_nodes = set()
            for node in list(G.nodes()):
                if initial_node_load[0][node] > initial_node_capacity[0][node]:
                    new_failed_nodes.add(node)

         # 如果没有新的节点崩溃，则结束模拟
            if not new_failed_nodes:
                break
            else:
                weight_total_matrix_P_j = np.zeros((project_number, 1))
                weight_node_matrix_P_j = np.zeros((project_number, project_number))
                for node in list(new_failed_nodes):
                    if initial_node_load[0][node] != 0:
                        nei_new_failed_nodes = G.neighbors(node)
                        kechuandi_nodes_for_new = list(set(nei_new_failed_nodes) - set(new_failed_nodes))
                        li_2 = []
                        if kechuandi_nodes_for_new != []:
                            for j in kechuandi_nodes_for_new:
                                weight_node_matrix_P_j[node][j] = initial_node_capacity[0][j] - initial_node_load[0][j]
                                if weight_node_matrix_P_j[node][j] < 0:
                                    print("---------------------------------------------出错了7---------------------------------------")
                                else:
                                    li_2.append(weight_node_matrix_P_j[node][j])
                            weight_total_matrix_P_j[node][0] = sum(li_2)
                        else:
                            continue
                    else:
                        continue

                for node in list(new_failed_nodes):
                    if initial_node_load[0][node] != 0:
                        nei_new_failed_nodes = G.neighbors(node)
                        kechuandi_nodes_for_new = list(set(nei_new_failed_nodes) - set(new_failed_nodes))
                        if kechuandi_nodes_for_new != []:
                            if weight_total_matrix_P_j[node][0] > 0:
                                for j in kechuandi_nodes_for_new:
                                    node_load_update_new = initial_node_load[0][j] + initial_node_load[0][node] * \
                                                           (weight_node_matrix_P_j[node][j] / weight_total_matrix_P_j[node][0])
                                    initial_node_load[0][j] = node_load_update_new
                            else:
                                for j in kechuandi_nodes_for_new:
                                    node_load_update_new = initial_node_load[0][j] + initial_node_load[0][node] / len(kechuandi_nodes_for_new)
                                    initial_node_load[0][j] = node_load_update_new
                        else:
                            continue
                    else:
                        continue
                G.remove_nodes_from(new_failed_nodes)
        print("-----------------------------------------项目层网络-----------------------------------------")
        print(G)

    # 计算能力实现值
    Realized_C = np.zeros((1, capacity_number)) # 为计算实际实现的能力值所构建的矩阵
    Expected_C = np.zeros((1, capacity_number)) # 为计算期望实现的能力值所构建的矩阵

    # 不考虑级联失效下能实现的最大值
    for i in range(capacity_number):
        li = []
        for nodes in list(G_star.nodes):
            li.append(P_S_C[nodes][i])
        Expected_C[0][i] = sum(li)

    for i in range(capacity_number):
        li_f = []
        for j in range(capacity_number):
            sum_number = Relationship_C[j][i] * Expected_C[0][j]
            li_f.append(sum_number)
        Expected_C[0][i] = Expected_C[0][i] + sum(li_f)
    CC_expected = np.sum(capacity_weight * Expected_C)

    # 级联失效后的实现的能力值
    for i in range(capacity_number):
        li = []
        for nodes in list(G.nodes):
            li.append(P_S_C[nodes][i])
        Realized_C[0][i] = sum(li)

    for i in range(capacity_number):
        li_r = []
        for j in range(capacity_number):
            sum_number = Relationship_C[j][i] * Realized_C[0][j]
            li_r.append(sum_number)
        Realized_C[0][i] = Realized_C[0][i] + sum(li_r)
    CC_realized = np.sum(capacity_weight * Realized_C)

    rate = CC_realized/CC_expected
    print(rate)
    cc_rate.append(rate)

    average = sum(cc_rate)/len(cc_rate)
    if IP == 1:
        rou_i = 0
    else:
        rou_i = average - cc_rate_average[-1]
        if rou_i <= 2 * 10**(-4):
            rou_i = 0
    cc_rate_average.append(average)
    rou.append(rou_i)

#绘图
x = range(0, iteration)
y = rou

print(cc_rate_average[-1])

plt.plot(x, y)

plt.savefig('PP_3.png')

#展示图形
plt.show()
