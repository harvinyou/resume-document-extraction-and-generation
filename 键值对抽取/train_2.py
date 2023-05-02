#  模型训练
import time
import os
import json
import random

import torch
import torch.nn as nn
import torch.optim as optim

from util import get_data_from_data_txt, get_word_to_ix, \
    prepare_sequence, get_score_by_label_pred, write_info_by_ix_plus
from data_process import get_str_from_pdf
from model import BiLSTM_CRF
from tqdm import tqdm

train_data = 'E://resumes//data//test_data.txt'

def get_score_by_model(model, train_json_path, pdf_root_dir):  # 评估模型
    pdf_path = os.listdir(pdf_root_dir)  # 用于返回指定的文件夹包含的文件或文件夹的名字的列表，用于能够遍历所有pdf数据
    random.shuffle(pdf_path)  # 打乱pdf文件的顺序
    path = pdf_path[:len(pdf_path)]
    word_to_ix = get_word_to_ix(train_data, 1)
    tag_to_ix = {'b-name': 0, 'i-name': 1, 'b-bir': 2, 'i-bir': 3, 'b-gend': 4, 'i-gend': 5,
                 'b-tel': 6, 'i-tel': 7, 'b-acad': 8, 'i-acad': 9, 'b-nati': 10, 'i-nati': 11,
                 'b-live': 12, 'i-live': 13, 'b-poli': 14, 'i-poli': 15, 'b-unv': 16, 'i-unv': 17,
                 'b-comp': 18, 'i-comp': 19, 'b-work': 20, 'i-work': 21, 'b-post': 22, 'i-post': 23,
                 'b-proj': 24, 'i-proj': 25, 'b-resp': 26, 'i-resp': 27, 'b-degr': 28, 'i-degr': 29,
                 'b-grti': 30, 'i-grti': 31, 'b-woti': 32, 'i-woti': 33, 'b-prti': 34, 'i-prti': 35,
                 'o': 36, '<start>': 37, '<stop>': 38, 'c-live': 39, 'c-proj': 40, 'c-woti': 41,
                 'c-post': 42, 'c-unv': 43, 'c-nati': 44, 'c-poli': 45, 'c-prti':46, 'c-comp': 47}
    ix_to_tag = {v: k for k, v in tag_to_ix.items()}

    pred_pdf_info = {}  # 显示经训练后数据标注结果的字典
    print('predicting...')
    for p in tqdm(path):
        if p.endswith('.pdf'):
            file_name = p[:-4]
            try:
                content = get_str_from_pdf(os.path.join(pdf_root_dir, p))  # pdf文件内容
                char_list = list(content)
                with torch.no_grad():  # 使用with torch.no_grad():表明当前计算不需要反向传播，使用之后，强制后边的内容不进行计算图的构建
                    precheck_sent = prepare_sequence(char_list, word_to_ix)
                    _, ix = model(precheck_sent)
                info = write_info_by_ix_plus(ix, content, ix_to_tag)
                pred_pdf_info[file_name] = info  # 字典key为filename，value为
            except Exception as e:
                if file_name not in pred_pdf_info:
                    pred_pdf_info[file_name] = {}
                print(e)
    print('predict OK!')
    # print(pred_pdf_info)
    with open(train_json_path, 'r',encoding='utf-8') as j:
        label_pdf_info = json.load(j)
    score = get_score_by_label_pred(label_pdf_info, pred_pdf_info)  # label_pdf_info是正确样本，pred_pdf_info是预测样本
    return score


def train_and_val():  # 训练和验证
    embedding_dim = 48 #原设置为100 由于标签一共有B\I\O\START\STOP...... 48个，所以embedding_dim为48 ?2倍
    hidden_dim = 100  # 这其实是BiLSTM的隐藏层的特征数量，因为是双向所以是2倍，单向为50
    best_model_save_path = 'E://resumes//model//best_model_test.pth'
    max_score = 0
    epoches = 3
    score = []

    # 训练数据
    training_data = get_data_from_data_txt(train_data)  # 训练数据 格式(['张','三'],['b-name','i-name'])

    # 处理数据集中句子的词，不重复地将句子中的词拿出来并标号
    # 设置一个word_to_ix存储句子中每一个单词
    word_to_ix = get_word_to_ix(train_data, 1)
    tag_to_ix = {'b-name': 0, 'i-name': 1, 'b-bir': 2, 'i-bir': 3, 'b-gend': 4, 'i-gend': 5,
                 'b-tel': 6, 'i-tel': 7, 'b-acad': 8, 'i-acad': 9, 'b-nati': 10, 'i-nati': 11,
                 'b-live': 12, 'i-live': 13, 'b-poli': 14, 'i-poli': 15, 'b-unv': 16, 'i-unv': 17,
                 'b-comp': 18, 'i-comp': 19, 'b-work': 20, 'i-work': 21, 'b-post': 22, 'i-post': 23,
                 'b-proj': 24, 'i-proj': 25, 'b-resp': 26, 'i-resp': 27, 'b-degr': 28, 'i-degr': 29,
                 'b-grti': 30, 'i-grti': 31, 'b-woti': 32, 'i-woti': 33, 'b-prti': 34, 'i-prti': 35,
                 'o': 36, '<start>': 37, '<stop>': 38, 'c-live': 39, 'c-proj': 40, 'c-woti': 41,
                 'c-post': 42, 'c-unv': 43, 'c-nati': 44, 'c-poli': 45, 'c-prti':46, 'c-comp': 47}

    # 将数据输入到 BiLSTM-CRF模型中
    model = BiLSTM_CRF(len(word_to_ix), tag_to_ix, embedding_dim, hidden_dim)
    optimizer = optim.Adam(model.parameters(), lr=0.01)  # 优化器


    for epoch in range(epoches):
        print("---------------------")
        print("running epoch : ", epoch)
        start_time = time.time()
        for sentence, tags in tqdm(training_data):
            model.zero_grad()
            # 输入
            sentence_in = prepare_sequence(sentence, word_to_ix)  #  将句子”w1 w2 w3”（词序列）[w1,w2,w3..]中的词w转化为在词表中的索引值 to_ix [w]
            targets = torch.tensor([tag_to_ix[t] for t in tags], dtype=torch.long)

            # 获取loss
            loss = model.neg_log_likelihood(sentence_in, targets)

            # 反向传播
            loss.backward()
            nn.utils.clip_grad_norm_(model.parameters(), 1)
            optimizer.step()

        val_json_path = 'E://resumes//data//valid_data_true.json'
        val_pdf_dir = 'E://resumes//data//valid_pdf_data'  # 验证集
        cur_epoch_score = get_score_by_model(model, val_json_path, val_pdf_dir)
        print('score', cur_epoch_score)
        score.append(cur_epoch_score)
        print('running time:', time.time() - start_time)
        if cur_epoch_score > max_score:  # 保存评分最高的模型
            max_score = cur_epoch_score
            torch.save({
                'model_state_dict': model.state_dict() # 保存整个模型
                # 'optimizer_state_dict': optimizer.state_dict(),  # 保存模型权重
                # 'epoch': epoch
            }, best_model_save_path)
            print('save best model successfully.')
        else:
            break

    # 绘图
    import matplotlib.pyplot as plt
    import numpy as np

    X = np.linspace(0, 1, epoches)  # X轴坐标数据
      # Y轴坐标数据
    plt.plot(X,score,color="red",linewidth=2)

    plt.figure(figsize=(8, 6))  # 定义图的大小
    plt.xlabel("epoch")  # X轴标签
    plt.ylabel("Score")  # Y轴坐标标签
    plt.title("The score of the model training")  # 曲线图的标题

    plt.plot(X, score)  # 绘制曲线图
    # 在ipython的交互环境中需要这句话才能显示出来
    plt.show()

if __name__ == '__main__':
    # train_all_data()
    train_and_val()