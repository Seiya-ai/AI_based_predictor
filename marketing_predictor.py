#!/usr/bin/env python
# coding: utf-8

# In[1]:

import logging

# ログファイルの設定
logging.basicConfig(filename='debug.log', level=logging.DEBUG)

# ログメッセージの出力
logging.debug('This is a debug message')
logging.info('This is an info message')
logging.warning('This is a warning message')
logging.error('This is an error message')
logging.critical('This is a critical message')

#ライブラリのインポート
#tkinter系ライブラリ
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

#ディープラーニング系ライブラリ
from sklearn.metrics import classification_report
from sklearn.model_selection import train_test_split
from tensorflow import keras
from tensorflow.keras import regularizers
from tensorflow.keras.layers import Dense, Activation
from tensorflow.keras.layers import Dropout

#その他必要なライブラリ
import codecs
import glob,os,re,time
import numpy as np
import os
import pandas as pd
import time
import traceback


# In[2]:


class DataX:
    def __init__(self):
        self.cnt_x = 0
        self.cnt_fnfe = 0
        self.cnt_str = 0

    #説明変数Xのファイルを選択するボタンの作成
    def set(self): 
        typ = [("エクセルファイル","*.xlsx")]
        file_path = filedialog.askopenfilename(filetypes = typ)
        file_name = os.path.split(file_path)[1]
        #以下エラーハンドリングと例外処理。
        #1.エクセルを読み込んで、意図しないエラーがあったときはメッセージを表示させる。
        #2.例外が発生するのでキャッチする。その際はログも表示させる。
        try:
            box1.delete(0, tk.END)
            df_x = pd.read_excel(file_path, index_col=0)
            self.df_x = df_x
            columns = df_x.columns.values.tolist()
            self.columns = columns
            self.cnt_x = 0
            self.cnt_fnfe = 0
            self.cnt_str = 0

            if df_x.index[0] != 0:
                messagebox.showerror('警告', "列名の中にnullが存在します。")
                self.cnt_x += 1
        
            for i in range(len(columns)):
                for row_is_str in df_x[columns[i]].apply(lambda x: isinstance(x, str)):
                    if row_is_str is True:
                       self.cnt_str += 1 
                
                if (self.cnt_str != 0) and df_x[columns[i]].isnull().values.any():
                    messagebox.showerror('警告', f'データの {columns[i]}列 に文字列とNullが含まれていました。データを見直して下さい。')
                    self.cnt_x += 1
                    self.cnt_str = 0                
                    
                elif (self.cnt_str != 0) :
                    messagebox.showerror('警告', f'データの {columns[i]}列 に文字列が含まれていました。該当列データを見直して下さい。')
                    self.cnt_x += 1
                    self.cnt_str = 0
                
                elif df_x[columns[i]].isnull().values.any():
                    messagebox.showerror('警告', f'データの {columns[i]}列 にNullが含まれていました。該当列データを見直して下さい。')
                    self.cnt_x += 1

        except UnicodeDecodeError as e:
            t_x = traceback.format_exception_only(type(e), e)
            messagebox.showerror('警告', "データの文字コードをutf-8に変更して下さい。\n"+t_x[0])
            self.cnt_x += 1
            
        except FileNotFoundError as e:
            t_x = traceback.format_exception_only(type(e), e)
            messagebox.showerror('警告', "ファイルが選択されていません。\n"+t_x[0])
            self.cnt_x += 1
            self.cnt_fnfe += 1
            #変数self.df_xを初期化しないと、ファイル選択していないのに前の情報を記憶しているため、初期化する必要がある
            try:
                del self.df_x
            except AttributeError as e:
                t_x = traceback.format_exception_only(type(e), e)
                if "DataX" in t_x[0]:
                    self.cnt_x += 1
                    
        if self.cnt_x == 0:
            box1.delete(0, tk.END)
            box1.insert(tk.END, f"「{file_name}」ファイルが正しく挿入されました。")  

        elif self.cnt_fnfe != 0:
            box1.delete(0, tk.END)
            box1.insert(tk.END, "ファイルが選択されていません。")
            self.cnt_fnfe = 0
            
        else:
            box1.delete(0, tk.END)
            box1.insert(tk.END, "【エラー】データを見直して下さい。")   
            self.cnt_x = 0


# In[3]:


class DataY:
    def __init__(self):
        self.cnt_y = 0
        self.cnt_fnfe_y = 0
        self.cnt_str_y = 0

    #目的変数Yのファイルを選択するボタンの作成
    def set(self): 
        typ = [("エクセルファイル","*.xlsx")]
        file_path_y = filedialog.askopenfilename(filetypes = typ)
        file_name = os.path.split(file_path_y)[1]
        #以下エラーハンドリングと例外処理。
        #1.エクセルを読み込んで、意図しないエラーがあったときはメッセージを表示させる。
        #2.例外が発生するのでキャッチする。その際はログも表示させる。
        #3.説明変数Xと行数が一致してるかや、目的変数Yの列名が複数存在していないかのチェックをおこなっている。
        try:
            box2.delete(0, tk.END)
            df_y = pd.read_excel(file_path_y, index_col=0)
            df_x = data_x.df_x
            self.df_y = df_y
            columns_y = self.df_y.columns.values.tolist()
            self.columns_y = columns_y
            self.cnt_y = 0
            self.cnt_fnfe_y = 0
            self.cnt_str_y = 0
            self.cnt_zero_one = 0
            
            if df_y.shape[0] != df_x.shape[0]:
                messagebox.showerror('警告', "説明変数データXと行数が一致していません。")
                self.cnt_y += 1
            
            if df_y.index[0] != 0:
                messagebox.showerror('警告', "列名の中にnullが存在します。")
                self.cnt_y += 1
                    
            if len(self.columns_y) != 1:
                messagebox.showerror("警告", "データの列数は一つである必要があります。")
                self.cnt_y += 1
                
            else:
                
                for row_is_str in df_y[columns_y[0]].apply(lambda x: isinstance(x, str)):
                    if row_is_str is True:
                        self.cnt_str_y += 1 
                        
                if (self.cnt_str_y != 0) and df_y[columns_y[0]].isnull().values.any():
                    messagebox.showerror("警告", "データに文字列とNullが含まれていました。データを見直して下さい。")
                    self.cnt_y += 1
                    self.cnt_str_y = 0                
                    
                elif (self.cnt_str_y != 0) :
                    messagebox.showerror('警告', "データに文字列が含まれていました。該当列データを見直して下さい。")
                    self.cnt_y += 1
                    self.cnt_str_y = 0
        
                elif df_y[columns_y[0]].isnull().values.any():
                    messagebox.showerror('警告', "データにNullが含まれていました。該当列データを見直して下さい。")
                    self.cnt_y += 1

                for number in df_y[columns_y[0]].values.tolist():
                    if number not in [0,1]:
                        self.cnt_zero_one += 1
                        
                if self.cnt_zero_one > 0:
                    messagebox.showerror('警告', "データに0か1以外の数値が含まれていました。データを見直して下さい。")
                    self.cnt_y += 1
                    
        except UnicodeDecodeError as e:
            t_y = traceback.format_exception_only(type(e), e)
            messagebox.showerror('警告', "データの文字コードをutf-8に変更して下さい。\n"+t_y[0])
            self.cnt_y += 1

        except FileNotFoundError as e:
            t_y = traceback.format_exception_only(type(e), e)
            messagebox.showerror('警告', "ファイルが選択されていません。\n"+t_y[0])
            self.cnt_y += 1
            self.cnt_fnfe_y += 1
            #変数self.df_yを初期化しないと、ファイル選択していないのに前の情報を記憶しているため、初期化する必要がある
            try:
                del self.df_y
            except AttributeError as e:
                t_y = traceback.format_exception_only(type(e), e)
                if "DataY" in t_y[0]:
                    self.cnt_y += 1
            except NameError as e:
                t_y = traceback.format_exception_only(type(e), e)
                if "DataY" in t_y[0]:
                    self.cnt_y += 1

        except NameError as e:
            t_y = traceback.format_exception_only(type(e), e)
            messagebox.showerror('警告', "説明変数データXが入力されていません。\n"+t_y[0])
            self.cnt_y += 1

        except AttributeError as e:
            t_y = traceback.format_exception_only(type(e), e)
            if "DataX" in t_y[0]:
                messagebox.showerror('警告', "説明変数Xが選択されていません\n"+t_y[0])
                self.cnt_y += 1
                    
        if self.cnt_y == 0:
            box2.delete(0, tk.END)
            box2.insert(tk.END, f"「{file_name}」ファイルが正しく挿入されました。")  

        elif self.cnt_fnfe_y != 0:
            box2.delete(0, tk.END)
            box2.insert(tk.END, "ファイルが選択されていません")
            self.cnt_fnfe_y = 0
            
        else:
            box2.delete(0, tk.END)
            box2.insert(tk.END, "【エラー】データを見直して下さい。")
            self.cnt_y = 0
            


# In[4]:


class DataX2:
    def __init__(self):
        self.cnt2 = 0
        self.cnt_fnfe_x2 = 0
        self.cnt_str2 = 0
        self.cnt_set_x2 = 0  
        self.columns_x2= 0
        self.df_x2 = 0

    #説明変数X2のファイルを選択するボタンの作成
    def set(self): 
        typ = [("エクセルファイル","*.xlsx")]
        file_path2 = filedialog.askopenfilename(filetypes = typ)
        file_name = os.path.split(file_path2)[1]
        #以下エラーハンドリングと例外処理。
        #1.エクセルを読み込んで、意図しないエラーがあったときはメッセージを表示させる。
        #2.例外が発生するのでキャッチする。その際はログも表示させる。
        #3.説明変数Xと行数が一致してるかや、説明変数Xや目的変数Yが選択されている状態かのチェックも行っている。
        try:
            box7.delete(0, tk.END)
            self.df_x2 = pd.read_excel(file_path2,index_col=0)
            self.columns_x2 = self.df_x2.columns.values.tolist()
            columns_x = data_x.columns
            df_x = data_x.df_x
            df_y = data_y.df_y
            self.cnt2 = 0                
            self.cnt_str2 = 0
            self.cnt_set_x2 = 0
            #self.cnt_fnfe_x2は何列も違う列名が存在した場合表示を一回にするためのカウンタ変数
            self.cnt_fnfe_x2 = 0 

            if self.df_x2.shape[0] != df_x.shape[0]:
                messagebox.showerror('警告', "説明変数データXと行数が一致しないか、データXが存在していません。")
                self.cnt2 += 1

            if self.df_x2.index[0] != 0:
                messagebox.showerror('警告', "列名の中にnullが存在します。")
                self.cnt2 += 1
                
            for i in range(len(self.columns_x2)):
                for row_is_str in self.df_x2[self.columns_x2[i]].apply(lambda x: isinstance(x, str)):
                    if row_is_str is True:
                       self.cnt_str2 += 1 
                
                if (self.cnt_str2 != 0) and self.df_x2[self.columns_x2[i]].isnull().values.any():
                    messagebox.showerror('警告', f'データの {self.columns_x2[i]}列 に文字列とNullが含まれていました。データを見直して下さい。')
                    self.cnt2 += 1
                    self.cnt_str2 = 0                
                    
                elif (self.cnt_str2 != 0) :
                    messagebox.showerror('警告', f'データの {self.columns_x2[i]}列 に文字列が含まれていました。該当列データを見直して下さい。')
                    self.cnt2 += 1
                    self.cnt_str2 = 0
                
                elif self.df_x2[self.columns_x2[i]].isnull().values.any():
                    messagebox.showerror('警告', f'データの {self.columns_x2[i]}列 にNullが含まれていました。該当列データを見直して下さい。')
                    self.cnt2 += 1
        
                elif set(columns_x) != set(self.columns_x2) and self.cnt_set_x2 == 0:
                    messagebox.showerror(tk.END, "学習用データと列名が一致しません。データを見直して下さい。")
                    self.cnt2 += 1
                    self.cnt_set_x2 += 1

        except UnicodeDecodeError as e:
            t_x2 = traceback.format_exception_only(type(e), e)
            messagebox.showerror('警告', "データの文字コードをutf-8に変更して下さい。\n"+t_x2[0])
            self.cnt2 += 1 
            
        except FileNotFoundError as e:
            t_x2 = traceback.format_exception_only(type(e), e)
            messagebox.showerror('警告', "ファイルが選択されていません。\n"+t_x2[0])
            self.cnt2 += 1
            self.cnt_fnfe_x2 += 1
            try:
                del self.df_x2
            except AttributeError as e:
                t_x2 = traceback.format_exception_only(type(e), e)
                if "DataX2" in t_x2[0]:
                    self.cnt2 += 1
            except NameError as e:
                t_x2 = traceback.format_exception_only(type(e), e)
                if "DataX2" in t_x2[0]:
                    self.cnt2 += 1
        
        except AttributeError as e:
            t_x2 = traceback.format_exception_only(type(e), e)
            if "DataY" in t_x2[0]:
                messagebox.showerror('警告', "目的変数Yが選択されていません\n"+t_x2[0])
                self.cnt2 += 1
            elif "DataX" in t_x2[0]:
                messagebox.showerror('警告', "説明変数Xが選択されていません\n"+t_x2[0])
                self.cnt2 += 1
            else:
                 messagebox.showerror('警告', "データに異常が発生しました。\n"+t_x2[0])
                 self.cnt2 += 1                 
            
        if self.cnt2 == 0:
            box7.delete(0, tk.END)
            box7.insert(tk.END, f"「{file_name}」ファイルが正しく挿入されました。")  

        elif self.cnt_fnfe_x2 != 0:
            box7.delete(0, tk.END)
            box7.insert(tk.END, "ファイルが選択されていません")
            
        else:
            box7.delete(0, tk.END)
            box7.insert(tk.END, "【エラー】データを見直して下さい。")


# In[18]:


class Gui:

    def __init__(self):
        self.data_with_model_new = 0
        self.y_test = 0
        self.x_test = 0
        self.trained_model = 0
        # self.learning_log = 0
        self.data_with_model = 0
        self.cnt_learningmodel = 0

    #説明変数Xと目的変数Yをもとに、最適な予測アルゴリズムを、ディープラーニングを用いて学習し、学習モデルを構築する。 
    #1.入力データ（説明変数Xや目的変数Y）に意図しないエラーがあったときはメッセージを表示させる。
    #2.例外が発生するのでキャッチする。その際はログも表示させる。
    def train_model(self):
    
        time.sleep(3)

        try:
            columns_y = data_y.columns_y
            df_x = data_x.df_x
            df_y = data_y.df_y
            
            if (len(columns_y) == 1) and (len(df_x)==len(df_y)):
                x = np.array(df_x[:])
                y = np.array(df_y[:])
                x_train, self.x_test, y_train, self.y_test = train_test_split(x, y, test_size=0.3, random_state=0, shuffle = True)
                x_train, x_valid, y_train, y_valid = train_test_split(x_train, y_train, test_size=0.3, random_state=0, shuffle = True)
            
                time.sleep(5)
                
                model = keras.Sequential()
                # 入力層
                model.add(Dense(len(df_x.columns), activation='relu', input_shape=(len(df_x.columns),), kernel_regularizer=regularizers.l2(0.01)))
                # 出力層
                model.add(Dense(1, activation='sigmoid'))
                # モデルの構築
                model.compile(optimizer = "adam", loss='binary_crossentropy', metrics=['accuracy'])
                self.trained_model = model
                
                time.sleep(3)
                
                # 学習の実施
                log = model.fit(x_train, y_train, epochs=1000, batch_size=32, verbose=True,
                            callbacks=[keras.callbacks.EarlyStopping(monitor='val_loss',
                                                                     min_delta=0, patience=20,
                                                                     verbose=1)],validation_data=(x_valid, y_valid))
                self.learning_log = log
            
                time.sleep(5)
                valid_loss, valid_acc = model.evaluate(x_valid, y_valid, verbose=0)
                
                box5.delete(0, tk.END)
                box5.insert(tk.END, "学習が完了しました")
                self.cnt_learningmodel += 1
                
            else:
                box5.delete(0, tk.END)
                box5.insert(tk.END, "問題が発生しました。行数があっているか、データが入力されているか確認してください")
                
        except AttributeError as e:
            t_gui_attribute_error = traceback.format_exception_only(type(e), e)
            if "DataX" in t_gui_attribute_error[0]:
                messagebox.showerror('警告', "説明変数Xに異常があります。データを見直して下さい。\n"+t_gui_attribute_error[0])
                box5.delete(0, tk.END)
                box5.insert(tk.END, "データを見直して下さい。") 
            elif "DataY" in t_gui_attribute_error[0]:
                messagebox.showerror('警告', "目的変数Yに異常があります。データを見直して下さい。\n"+t_gui_attribute_error[0])
                box5.delete(0, tk.END)
                box5.insert(tk.END, "データを見直して下さい。") 
            else:
                messagebox.showerror('警告', "異常があります。データを見直して下さい。\n"+t_gui_attribute_error[0])
                box5.delete(0, tk.END)
                box5.insert(tk.END, "データを見直して下さい。") 

    #学習したデータ（説明変数Xと目的変数Y）で予測モデルの構築を行い、実際にどれくらい予測が正解したのかを表示させるボタンの作成                
    def show_accuracy_score(self):
        y_pred = self.trained_model.predict(self.x_test)
        # 二値分類は予測結果の確率が0.5以下なら0,
        # それより大きければ1となる計算で求める
        time.sleep(2)
        y_pred_probability = (y_pred > 0.5).astype("int32")
        y_pred_ = y_pred_probability.reshape(-1)
        cr = classification_report(self.y_test, y_pred_)
        
        if self.cnt_learningmodel > 0:
            box6.delete('1.0', tk.END)
            box6.insert('1.0',cr) 
        else:
            box6.insert('1.0',cr) 

    

    #構築した予測モデルを、分析したいデータ（説明変数X2）に適用する。    
    def analyze_newdata(self):
        if set(data_x.columns) == set(data_x2.columns_x2):
            x_test2 = np.array(data_x2.df_x2[:])
            y_pred2 = self.trained_model.predict(x_test2)
            # 二値分類は予測結果の確率が0.5以下なら0,
            # それより大きければ1となる計算で求める    
            time.sleep(2)
            nd_data = np.concatenate([x_test2, y_pred2], axis=1)
            self.data_with_model = pd.DataFrame(nd_data)
            time.sleep(2)
            box9.delete(0, tk.END)
            box9.insert(tk.END,"予測が完了しました")
        else:
            box9.delete(0, tk.END)
            box9.insert(tk.END,"学習データと列名が一致しません。データを見直して下さい。")
        
    #確率のレベルに応じて色のグラデーションをつける
    def color_highlight(self,x):
        if x[len(x)-1]>0.8:        
            return['background-color: #c30010']*len(x)
        elif x[len(x)-1]>0.6:
            return['background-color: #f45d75']*len(x)
        elif x[len(x)-1]>0.4:
            return['background-color: #f7b4bb']*len(x)
        elif x[len(x)-1]>0.2:
            return['background-color: #f7d7dc']*len(x)
        else:
            return['background-color: white']*len(x)

        # if x[len(x)-1]>0.8:        
        #     return ['background-color: #c30010']
        # elif x[len(x)-1]>0.6:
        #     return ['background-color: #f45d75']
        # elif x[len(x)-1]>0.4:
        #     return ['background-color: #f7b4bb']
        # elif x[len(x)-1]>0.1:
        #     return ['background-color: #f7d7dc']
        # else:
        #     return ['background-color: white']

    #色を背景色に設定する        
    def highlight_rows(self):
        data_with_model = self.data_with_model
        self.data_with_model_new = data_with_model.style.apply(self.color_highlight, axis =1) 
        time.sleep(2)
        box10.delete(0, tk.END)
        box10.insert(tk.END,"データの成型が完了しました") 

    #エクセルに出力する        
    def to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            data = self.data_with_model_new
            data.to_excel(file_path,"analyzed_data.xlsx")
            time.sleep(2)
            box11.delete(0, tk.END)
            box11.insert(tk.END,"エクセルファイルの出力が完了しました")    


# In[19]:


#GUIウィンドウやボタンやメッセージ欄のレイアウトを作る
#実際にプログラムを起動させる(上記のクラスが実行されないと起動しない)

data_x = DataX()
data_y = DataY()
data_x2 = DataX2()
gui = Gui()

main_win = tk.Tk()
main_win.title("簡易ディープラーニングツール")
main_win.geometry("600x500")
box1 = tk.Entry(width=40, text = "xlsxファイルのフルパス")
box1.place(x=20,y=30)
box1_label = tk.Label(text = "学習用データ(説明変数)")
box1_label.place(x=270,y=30)
box1.insert(tk.END, "ファイルを選択してください")
set_button = tk.Button(text = "ファイルを選択(xlsxファイル)", command = data_x.set)
set_button.place(x = 419, y = 27)

box2 = tk.Entry(width=40)
box2.place(x=20,y=70)
box2_label = tk.Label(text = "学習用データ(目的変数)")
box2_label.place(x=270,y=70)
box2.insert(tk.END, "ファイルを選択してください")
set_button = tk.Button(text = "ファイルを選択(xlsxファイル)", command = data_y.set)
set_button.place(x = 419, y = 67)

box5 = tk.Entry(width=40)
box5.place(x=20,y=110)
box5_label = tk.Label(text = "完了ステータス")
box5_label.place(x=270,y=110)
set_button = tk.Button(text = "機械学習モデルを構築", command= gui.train_model)
set_button.place(x = 419, y = 107)

box6 = tk.Text()
box6.pack()
box6.place(x=20,y=160, width=380,height=120)
set_button = tk.Button(text = "モデルの予測精度を表示", command = gui.show_accuracy_score)
set_button.place(x = 419, y = 197)

box7 = tk.Entry(width=40)
box7.place(x=20,y=310)
box7_label = tk.Label(text = "分析したいデータ(説明変数)")
box7_label.place(x=270,y=310)
box7.insert(tk.END, "ファイルを選択してください")
set_button = tk.Button(text = "新しいファイルを選択(xlsxファイル)", command = data_x2.set)
set_button.place(x = 419, y = 307)

box9 = tk.Entry(width=40)
box9.place(x=20,y=350)
box9_label = tk.Label()
box9_label.place(x=270,y=350)
set_button = tk.Button(text = "新しいデータを予測", command = gui.analyze_newdata)
set_button.place(x = 419, y = 347)

box10 = tk.Entry(width=40)
box10.place(x=20,y=390)
box10_label = tk.Label()
box10_label.place(x=270,y=390)
set_button = tk.Button(text = "データを成型", command = gui.highlight_rows)
set_button.place(x = 419, y = 387)

box11 = tk.Entry(width=40)
box11.place(x=20,y=430)
box11_label = tk.Label()
box11_label.place(x=270,y=430)
set_button = tk.Button(text = "エクセルファイルを出力", command = gui.to_excel)
set_button.place(x = 419, y = 427)

main_win.mainloop()


# In[17]:


pd.read_excel("analyzed_data.xlsx")


# In[ ]:


gui.data_with_model.style.apply(gui.color_highlight, axis =1) 


# In[ ]:


x[len(x.columns)-1]


# In[ ]:


import pandas as pd

# サンプルデータフレームを作成
data = {'A': [0.1, 0.7, 0.3, 0.9], 'B': [0.5, 0.2, 0.8, 0.4]}
df = pd.DataFrame(data)

# カスタムのスタイリング関数
def highlight_column(col):
    # 基準となる閾値を設定
    threshold = 0.5
    # スタイリング情報を保持するリスト
    styles = []
    
    for val in col:
        if val > threshold:
            styles.append('background-color: red')  # 条件を満たす場合の背景色
        else:
            styles.append('background-color: green')  # 条件を満たさない場合の背景色
    
    return styles

# 特定の列にカスタムのスタイリング関数を適用
styled_df = df.style.apply(highlight_column, subset=['A'], axis=0)

# スタイリングが適用されたデータフレームを表示
styled_df


# %%
