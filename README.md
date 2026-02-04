# Slb_Relation_Build
### 功能说明

1. 通过TextFSM提取CLI回显中的关键字段，通过pd进行关联关系处理与整合输出。
2. 本项目的核心在于通过TextFSM的提取内容，因不同LB版本的CLI回显不一致，可能会导致TextFSM解析存在异常。

3. TextFSM模板本质是正则匹配，通过本地log文件进行数据提取会存在不可控因素，虽然代码中添加了数据数量校验逻辑，但仍然无法确保数据提取关联100%准确

### 本地log文件要求

**迪普：将show run单独放置于一个log文件，并且命名需要以-conf结尾，剩余的放在一个log文件，并且命名需要以-slb结尾，详细格式请参考Log文件夹**

```bash
show run

show slb virtual-service status
show slb pool status
show slb member status
```

**信安：**

```bash
show tech
```

**弘积：**

```bash
show slb virtual-address
show slb pool
show run
```

**本项目仅提供实现思路，欢迎学习交流。**