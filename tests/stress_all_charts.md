# 图表类型全览 — 压力测试

## 柱状图（Column）

```chart
{"type": "column", "categories": ["Q1","Q2","Q3","Q4"], "values": [120,190,150,220], "series_name": "营收"}
```

## 条形图（Bar）

```chart
{"type": "bar", "categories": ["产品A","产品B","产品C","产品D","产品E"], "values": [85,72,93,68,77], "series_name": "用户评分"}
```

## 折线图（Line）

```chart
{"type": "line", "categories": ["1月","2月","3月","4月","5月","6月"], "values": [45,52,48,61,58,67], "series_name": "活跃用户数（万）"}
```

## 饼图（Pie）

```chart
{"type": "pie", "categories": ["iOS","Android","Web","Desktop"], "values": [35,42,15,8], "series_name": "平台分布"}
```

## 环形图（Doughnut）

```chart
{"type": "doughnut", "categories": ["北京","上海","深圳","杭州","成都","其他"], "values": [28,22,18,12,8,12], "series_name": "城市分布"}
```

## 面积图（Area）

```chart
{"type": "area", "categories": ["2019","2020","2021","2022","2023","2024","2025"], "values": [10,25,45,80,130,200,280], "series_name": "累计用户数"}
```

## 雷达图（Radar）

```chart
{"type": "radar", "categories": ["性能","安全","易用","扩展","文档","社区"], "values": [90,85,75,92,70,88], "series_name": "产品能力评估"}
```

## 瀑布图（Waterfall）

```chart
{"type": "waterfall", "categories": ["基线","效率","成本","增长","净收益"], "values": [0,120,80,150,350], "series_name": "利润拆解"}
```

## 漏斗图（Funnel）

```chart
{"type": "funnel", "categories": ["曝光","点击","访问","注册","付费","复购"], "values": [100000,25000,8000,3000,800,240], "series_name": "转化漏斗"}
```

## 甘特图（Gantt）

```chart
{"type": "gantt", "categories": ["需求","设计","开发","测试","上线","运营"], "values": [10,20,45,20,5,30], "series_name": "项目周期（天）"}
```
EOF