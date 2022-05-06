## fudan_tools

### daily_food_check.py
> 此项目用于解决复旦大学研究生宿舍订饭的时候填写表格错误问题，例如检查支付宝截图中支付的金额和表格填写的金额是否一致等
-  1 必须使用.xls后缀的表格文件，如果是.xlsx后缀的文件，请打开并另存为.xls文件，样例中是data.xls文件
-  2 文件中必须有以下列：序号、所在学院（必填）、姓名（必填）、早餐、午餐、午餐白米饭、晚餐、晚餐白米饭、支付总金额、支付宝付款截图上传
-  3 冒号后面的都为列名，必填也是列名中的一部分
-  4 支付总金额为表格中填写的那一列
-  5 run：程序会自动将错误结果保存下来为result.xls文件
-  6 运行前，请修改文件开头的最大线程数：max_workers 为合适的数目，建议<10, 且第一次运行请将其设置为1（下载必要包）。并且修改main函数开头导入文件的路径
-  7 注：需要GPU加速，如果硬件设备不支持，可以使用google-colab或天池等GPU开放平台进行运行
