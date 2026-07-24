# NNRoad 公共账号库 SQL（按你的确认定制）

- 公共库名：`nnroad_auth`（你写的 nnroad_aut，按 auth 理解）
- 初始数据来源：`nnroad_portal`
- 先改代码：Office（`nnroad-dev`）
- Office 业务主库名：`nnroad-office`（注意带连字符）

## 执行顺序（在能连上 portal/office 的 MySQL 上）

1. `01_sys_menu_app_code.sql`
2. `02_nnroad_auth_init_from_portal.sql`
3. `03_nnroad_auth_merge_office_menus.sql`

执行前请备份相关库。

## 注意

- 若 Office 有 portal 没有的管理员账号，在 02 脚本末尾可选块取消注释合并 `nnroad-office.sys_user`。
- 角色菜单合并依赖 `role_id` 一致；Office 独立角色需额外处理。
- 代码切 AUTH 后，yml 增加 `spring.datasource.druid.auth` 指向 `nnroad_auth`。
