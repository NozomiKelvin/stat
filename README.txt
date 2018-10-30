Hello World!

0、双击执行package/start.bat即可运行程序！

1、package目录结构说明
	|--- conf
	| |--- excel
	| | |--- 存放赋值，企业透视图Excel
	| |--- stat.properties
	|  
	|--- jre（不用管，千万别删）
	|
	|--- lib（不用管，千万别删）
	|
	|--- start.bat（双击运行）
	|
	|--- stat-1.0.0.jar（不用管，千万别删）

2、package/conf/stat.properties配置文件

a、输入文件相关配置
# 主数据来源，文件名
stat.movie.main-data.file-name=企业透视图列子.xlsx
# 主数据来源，工作簿名。可配置多个，半角逗号分隔（例如：A,B,C）
stat.movie.main-data.sheet-names=例子,例子2

# 关系数据来源，文件名
stat.movie.relation-data.file-name=赋值.xlsx
# 关系数据来源，工作簿名
stat.movie.relation-data.sheet-name=赋值表

b、输出文件相关配置
# 最终生成统计数据的工作表类型（xlsx或者xls，默认xlsx）
stat.movie.output-data.suffix=xlsx

3、src目录不解释了

4、Enjoy it!
