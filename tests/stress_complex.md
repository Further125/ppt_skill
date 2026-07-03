# 复杂结构压力测试

## 三级树状图

```tree
{
  "root": "公司战略",
  "children": [
    {
      "label": "产品战略",
      "children": [
        {"label": "核心产品升级"},
        {"label": "新产品孵化"},
        {"label": "产品矩阵扩展"}
      ]
    },
    {
      "label": "市场战略",
      "children": [
        {"label": "国内市场深耕"},
        {"label": "海外市场拓展"},
        {"label": "渠道合作伙伴"}
      ]
    },
    {
      "label": "技术战略",
      "children": [
        {"label": "AI能力平台化"},
        {"label": "云原生架构"},
        {"label": "数据安全合规"}
      ]
    },
    {
      "label": "人才战略",
      "children": [
        {"label": "顶尖人才引进"},
        {"label": "内部培养体系"},
        {"label": "激励机制优化"}
      ]
    }
  ]
}
```

## 大型团队（10人）

```team
{
  "members": [
    {"name": "Member H", "role": "CEO", "desc": "企业管理专家"},
    {"name": "Member A", "role": "CTO", "desc": "AI技术专家"},
    {"name": "Member B", "role": "CPO", "desc": "15年产品经验"},
    {"name": "Member C", "role": "设计总监", "desc": "资深体验设计师"},
    {"name": "Member D", "role": "市场VP", "desc": "品牌增长专家"},
    {"name": "Member E", "role": "销售VP", "desc": " enterprise销售负责人"},
    {"name": "Member I", "role": "架构师", "desc": "分布式系统专家"},
    {"name": "Member F", "role": "数据科学家", "desc": "算法竞赛金牌"},
    {"name": "Member G", "role": "工程总监", "desc": "DevOps布道者"},
    {"name": "Member J", "role": "HR总监", "desc": "组织发展专家"}
  ]
}
```

## 多阶段流程图

```chart
{"type": "process", "categories": ["需求分析","原型设计","技术评审","开发实现","代码审查","测试验证","灰度发布","全量上线","监控运维","复盘优化"], "values": [5,3,2,15,3,10,5,2,30,3], "series_name": "各阶段耗时（天）"}
```

## 甘特图

```chart
{"type": "gantt", "categories": ["需求","设计","前端","后端","测试","部署","验收"], "values": [7,10,20,25,15,3,5], "series_name": "计划工期（天）"}
```

## 漏斗图

```chart
{"type": "funnel", "categories": ["曝光","点击","落地页","注册","激活","留存","付费"], "values": [1000000,150000,75000,30000,12000,4800,1440], "series_name": "全链路转化"}
```

## 数据高亮

```highlight
{"big_number": "¥2.4亿", "label": "本季度营收创新高，同比增长340%"}
```

## 数据高亮2

```highlight
{"big_number": "99.99%", "label": "系统可用性，全年仅停机52分钟"}
```

## 数据高亮3

```highlight
{"big_number": "<50ms", "label": "API平均响应时间，P99<200ms"}
```
EOF