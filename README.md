# SEO-Data-Analysis


**一、seo工具：在一组数据中，聚合相似的句子**

**FrequencyOfOccurrenceSortAndWordNum_pop_cleanData_findsentence.py**

目的：excel一列中文句子中，聚合相似的句子，找到共同的关键词/关键句

处理顺序：

1，jieba分词，
2，然后n-gram拼成 1-n 个词组合的字符串，
3，去掉需要忽略的词，
4，根据词的长度（len_rate）和在句子中出现的频率(num_rate)作为权重，计算全部数据中每个词的权重，
5，找到每条数据中权重最高的3个词，作为这条数据中的关键词，并显示其权重
