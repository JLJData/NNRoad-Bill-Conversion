-- ---------------------------------------------------------------------------
-- 将 Portal + Office 菜单合并到 hrone_auth.sys_menu（单表 + app_code 隔离）
--
-- 前提：
--   1. 已执行 sys_menu_app_code.sql（hrone_portal 或 hrone_auth 先有 app_code 列）
--   2. 已执行 hrone_auth_init_from_portal.sql（Portal 菜单已在 hrone_auth，menu_id 不变）
--   3. Office 原始菜单在库 hrone_office.sys_menu（若从未初始化，见文末「方案 B」）
--
-- 策略：
--   - Portal：app_code = 'portal'，保留原 menu_id（角色、业务表都依赖它）
--   - Office：app_code = 'office'，menu_id / parent_id 整体 +100000，避免主键冲突
--   - sys_role_menu：Office 侧 menu_id 同步偏移后写入 hrone_auth
--
-- 执行前备份：
--   mysqldump -u root -p hrone_auth > hrone_auth_backup.sql
-- ---------------------------------------------------------------------------

USE `hrone_auth`;

-- 1) Portal 菜单打标（从 Portal 迁入的默认已是 portal，再执行一次无害）
UPDATE `sys_menu` SET `app_code` = 'portal' WHERE `app_code` IS NULL OR `app_code` = '' OR `app_code` = 'portal';

-- 2) Office 菜单 ID 偏移量（与 Portal 的 menu_id 区间错开即可，可按需改大）
SET @OFFICE_MENU_OFFSET = 100000;

-- 3) 从 hrone_office 合并菜单（若 Office 库名不同请改库名）
--    已存在同 app_code+menu_id 则跳过，脚本可重复执行
INSERT INTO `hrone_auth`.`sys_menu` (
    `menu_id`, `app_code`, `menu_name`, `parent_id`, `order_num`, `path`, `component`, `query`,
    `route_name`, `is_frame`, `is_cache`, `menu_type`, `visible`, `status`, `perms`, `icon`,
    `create_by`, `create_time`, `update_by`, `update_time`, `remark`
)
SELECT
    o.`menu_id` + @OFFICE_MENU_OFFSET,
    'office',
    o.`menu_name`,
    CASE WHEN o.`parent_id` = 0 THEN 0 ELSE o.`parent_id` + @OFFICE_MENU_OFFSET END,
    o.`order_num`, o.`path`, o.`component`, o.`query`, o.`route_name`,
    o.`is_frame`, o.`is_cache`, o.`menu_type`, o.`visible`, o.`status`, o.`perms`, o.`icon`,
    o.`create_by`, o.`create_time`, o.`update_by`, o.`update_time`, o.`remark`
FROM `hrone_office`.`sys_menu` o
WHERE NOT EXISTS (
    SELECT 1 FROM `hrone_auth`.`sys_menu` t
    WHERE t.`menu_id` = o.`menu_id` + @OFFICE_MENU_OFFSET
);

-- 4) Office 角色-菜单关联（menu_id 同步偏移）
--    注意：role_id 与 Portal 共用同一套 sys_role（来自 Portal 初始化）
--    若 Office 有独立角色体系，需先合并 sys_role 或新建 office 专用角色
INSERT IGNORE INTO `hrone_auth`.`sys_role_menu` (`role_id`, `menu_id`)
SELECT
    rm.`role_id`,
    rm.`menu_id` + @OFFICE_MENU_OFFSET
FROM `hrone_office`.`sys_role_menu` rm
INNER JOIN `hrone_office`.`sys_menu` m ON m.`menu_id` = rm.`menu_id`
WHERE EXISTS (
    SELECT 1 FROM `hrone_auth`.`sys_menu` t
    WHERE t.`menu_id` = rm.`menu_id` + @OFFICE_MENU_OFFSET AND t.`app_code` = 'office'
);

-- 5) 给「超级管理员」角色补全 Office 全部菜单（Office 库无 role_menu 或需一键授权时用）
--    超级管理员 role_id 一般为 1，请按实际修改
-- INSERT IGNORE INTO `hrone_auth`.`sys_role_menu` (`role_id`, `menu_id`)
-- SELECT 1, `menu_id` FROM `hrone_auth`.`sys_menu` WHERE `app_code` = 'office';

-- 6) 校验
SELECT `app_code`, COUNT(*) AS menu_cnt FROM `sys_menu` GROUP BY `app_code`;

SELECT r.`role_id`, r.`role_name`, m.`app_code`, COUNT(*) AS menu_cnt
FROM `sys_role_menu` rm
JOIN `sys_role` r ON r.`role_id` = rm.`role_id`
JOIN `sys_menu` m ON m.`menu_id` = rm.`menu_id`
GROUP BY r.`role_id`, r.`role_name`, m.`app_code`
ORDER BY r.`role_id`, m.`app_code`;

-- ---------------------------------------------------------------------------
-- 方案 B：hrone_office 里没有菜单（从未跑过 RuoYi 初始化）
-- 直接在 hrone_auth 插入 Office 种子，并设 app_code='office'，ID 仍建议 >= 100000：
--
-- SET @O = 100000;
-- INSERT INTO sys_menu (menu_id, app_code, menu_name, parent_id, ...) VALUES
--   (@O+1, 'office', '系统管理', 0, ...),
--   (@O+100, 'office', '用户管理', @O+1, ...);
-- 或把 ry_20260417.sql / system_home_app.sql 的 INSERT 改写成带 app_code 和偏移 ID。
-- ---------------------------------------------------------------------------
