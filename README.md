### Python处理Excel数据
___
> 使用python处理基本的Excel数据,以某个Excel文件为主模板,向其追加数据,暂时只支持处理单Sheet的Excel文件

#### 数据区
> data_dir = r"./data"

此属性用于指定对应的数据区,程序将从此目录下获得Excel中全部的数据,并最终转换为dataframe数据

#### 空行
> dataframe.loc[len(dataframe.index)] = ""

程序会自动在一个文件数据的尾部添加空行,以便于区分不同文件中的数据,如果您不需要,可以此行代码删除

#### 模板

模板是一个原型文件,这个文件通常放在model目录下面,所有数据都会按照此文件的格式来添加,并且不会破坏此文件原有的数据

#### 追加和输出
> excel_model = r'model/doc1.xls' # 模板路径   
> new_file = r'./new_file/doc1.xls' # 输出文件路径  
> write_excel_xls_append(excel_model, new_file, df)

使用write_excel_xls_append()函数能够完成数据的追加和文件的输出,这个函数需要三个参数,模板,输出文件的的路径,以及
对应的dataframe数据

#### 后续

后续将更新Excel格式处理模块