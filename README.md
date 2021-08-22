# MST_Prim
采用prim算法实现最小生成树，使用示例：节点为2017年河南省各城市，边为各城市之间的距离。 

输入为【距离_2017.xlsx】，为河南及其邻省的部分城市有接壤的距离。
![图片](https://user-images.githubusercontent.com/47874610/130345811-9dc057ed-83e6-49cc-a699-851c5e5cd5cf.png)

通过Prim算法，实现最小生成树【2017out.docx】

![gogogo](https://user-images.githubusercontent.com/47874610/130346002-c954d149-123d-482a-a3a4-e857484e0e0f.png)


之后进行剪枝：

①边长小于15，直接不剪。

②2个步长取样本边，若该边大于样本边平均值2倍，或该边在样本边的3倍标准差之外，则为不连续边，剪去。



