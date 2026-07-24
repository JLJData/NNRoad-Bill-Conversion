# 一期公共账号统一 · 同事方案捕捉（切回自己分支后按此改）

> 捕捉来源：Portal `HROne_portal_boot` 分支 `dev`（`ee9938d` 等）；Office `hrone-office-abp` 历史提交 `a3c789c` / `1e550f3`（SQL 后被 `5caa771` 删除，已从 git 还原到本目录）。
>
> **本期不做站内信 / `hrone_common`**，只借鉴公共账号 `hrone_auth`。

---

## 1. 目标架构（一期）

| 库 | 用途 |
|----|------|
| `hrone_auth` | 公共账号：用户、角色、菜单、部门、岗位、登录日志、密码重置等 |
| `hrone_office` / `hrone_portal` | 各端业务主库（先不拆业务主从同步） |

会话隔离（建议一并做）：

| 项 | Portal | Office |
|----|--------|--------|
| `hrone.app-code` | `portal` | `office` |
| JWT `token.secret` | 各自独立 | 各自独立 |
| Redis database | 建议 2 | 建议 1（prod 可 9） |
| Cache key 前缀 | `portal:` | `office:` |

---

## 2. 代码改造清单

1. yml：增加 `spring.datasource.druid.auth`；配置 `hrone.app-code`；独立 JWT / Redis db  
2. `DataSourceType` 增加 `AUTH`（COMMON 可暂不加）  
3. `DruidConfig` 注册 `authDataSource`（有 `auth.url` 才创建）  
4. `DynamicDataSourceContextHolder` 必须用 **栈 push/pop**（支持 AUTH 内再调 MASTER）  
5. `CacheConstants` 加应用前缀  
6. `AppConstants`：`office` / `portal`；`SysMenu.appCode`；菜单查询全带 `app_code`  
7. 类上 `@DataSource(DataSourceType.AUTH)`：  
   - SysUser / SysRole / SysMenu / SysDept / SysPost / SysLogininfor  
   - Portal 还有 SysPasswordReset  
8. 业务表（客户绑定、账单等）留 MASTER，不要迁 auth  
9. Office 改角色菜单时：勿删另一 app 的 `sys_role_menu`（保留 exclude 逻辑）  
10. Portal 客户菜单：MASTER 查客户角色 ID → AUTH 按 roleIds + app_code 查菜单  

---

## 3. `hrone_auth` 表清单与字段

### 3.1 表清单

| 表 | 说明 | 特殊 |
|----|------|------|
| `sys_user` | 公共用户 | `user_type`：00 超管 / 01 员工 / 10 企业客户 |
| `sys_role` | 公共角色 | |
| `sys_menu` | 多应用菜单 | **`app_code`：office / portal** |
| `sys_dept` | 部门 | |
| `sys_post` | 岗位 | |
| `sys_user_role` | 用户-角色 | |
| `sys_role_menu` | 角色-菜单 | Office 菜单 ID 常 +100000 |
| `sys_role_dept` | 角色-部门 | |
| `sys_user_post` | 用户-岗位 | |
| `sys_logininfor` | 登录日志 | |
| `sys_password_reset` | 密码重置 | Portal 常用 |
| `sys_role_user_type` | 角色适用用户类型 | **init 脚本未拷，代码若用到需单独补** |

### 3.2 `sys_user` 字段

`user_id, dept_id, user_name, nick_name, email, phonenumber, sex, user_type, avatar, password, status, del_flag, login_ip, login_date, pwd_update_date, create_by, create_time, update_by, update_time, remark`

### 3.3 `sys_menu` 关键字段

| 列 | 说明 |
|----|------|
| `menu_id` | PK；Office 合并后建议 +100000 |
| **`app_code`** | varchar(32) NOT NULL DEFAULT `'portal'`；`office` / `portal` |
| 其余 | menu_name, parent_id, order_num, path, component, query, route_name, is_frame, is_cache, menu_type, visible, status, perms, icon, create_*, update_*, remark |

### 3.4 关联表字段

| 表 | 字段 |
|----|------|
| `sys_user_role` | user_id, role_id |
| `sys_role_menu` | role_id, menu_id |
| `sys_role_dept` | role_id, dept_id |
| `sys_user_post` | user_id, post_id |
| `sys_role_user_type` | role_id, user_type |

### 3.5 仍在业务库、靠 user_id 关联 auth

| 表（Portal 示例） | 作用 |
|------------------|------|
| `portal_user_client` | 用户↔客户 |
| `portal_user_client_role` | 客户维度角色（role 定义在 auth） |

迁移 auth 时 **必须保留 `user_id`**，否则业务绑定断掉。

---

## 4. SQL 执行顺序（本目录已备脚本）

1. `sys_menu_app_code.sql` — 给 `sys_menu` 加 `app_code`（在源库或 auth 执行，视迁移步骤）  
2. `hrone_auth_init_from_portal.sql` — 建库并从 `hrone_portal` 拷贝认证表  
3. `hrone_auth_merge_menus.sql` — 合并 Office 菜单（ID+100000，`app_code=office`）  
4. 如代码用到：单独 `CREATE … LIKE` + `INSERT` **`sys_role_user_type`**

---

## 5. 配置片段参考（dev）

```yaml
hrone:
  app-code: office   # Portal 写成 portal

spring:
  datasource:
    druid:
      master:
        url: jdbc:mysql://.../hrone_office   # Portal 则为 hrone_portal
      auth:
        url: jdbc:mysql://.../hrone_auth
        username: root
        password: '***'
  data:
    redis:
      database: 1   # Portal 用 2

token:
  secret: hrone-office-jwt-secret-key-2026  # Portal 用另一把
```

---

## 6. 切回自己分支后要我做的事

请说「切回来了，按捕捉改公共账号」，我将按本备忘改：

- Office / Portal 多数据源 AUTH  
- 栈式 ContextHolder  
- app-code / 菜单过滤  
- （按需）执行或校对 SQL  
- **不接站内信 / common**

---

## 7. 一句话

**一期 = 账号权限进共享库 `hrone_auth`，菜单用 `app_code` 区分两端；业务库先不拆主从同步。**
