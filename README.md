# python-add-tracking-no

# 安装方法
1. 将压缩包解压即可

# 使用方法
1. 将_ship.csv和_express.xlsx结尾的两个数据文件复制到本目录
2. 双击exe程序运行
3. 看到*_ship.csv filename?提示后，输入_ship.csv文件的文件名，输入完毕回车
4. 看到*_express.xlsx filename?提示后，输入_express.xlsx文件的文件名，输入完毕回车
5. 程序运行完毕出现如下几个文件
	*_express.xlsx为填好Tracking NO的文件
	copy_*_express.xlsx为原*_express.xlsx的备份文件
	only_in_*_express.txt为只出现在*_express.xlsx中没有找到Tracking NO的订单号
	only_in_*_ship.txt为只出现在*_ship.csv中没有用到的订单号
6. **提示：这四个文件运行完毕后要移出目录才能进行下一次运行，否则数据会混乱**
