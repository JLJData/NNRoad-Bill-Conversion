-- ---------------------------------------------------------------------------
-- 03 将 Office 菜单合并进 nnroad_auth（ID +100000，app_code=office）
-- 前提：已执行 01、02
-- 执行前备份：mysqldump -u root -p nnroad_auth > nnroad_auth_backup.sql
-- ---------------------------------------------------------------------------

USE `nnroad_auth`;

UPDATE `sys_menu` SET `app_code` = 'portal' WHERE `app_code` IS NULL OR `app_code` = '' OR `app_code` = 'portal';

SET @OFFICE_MENU_OFFSET = 100000;

INSERT INTO `nnroad_auth`.`sys_menu` (
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
FROM `nnroad-office`.`sys_menu` o
WHERE NOT EXISTS (
    SELECT 1 FROM `nnroad_auth`.`sys_menu` t
    WHERE t.`menu_id` = o.`menu_id` + @OFFICE_MENU_OFFSET
);

-- Office 角色-菜单（menu_id 同步偏移）
-- 注意：role_id 当前来自 portal 初始化；若 Office 有独立角色，需另行合并 sys_role
INSERT IGNORE INTO `nnroad_auth`.`sys_role_menu` (`role_id`, `menu_id`)
SELECT
    rm.`role_id`,
    rm.`menu_id` + @OFFICE_MENU_OFFSET
FROM `nnroad-office`.`sys_role_menu` rm
INNER JOIN `nnroad-office`.`sys_menu` m ON m.`menu_id` = rm.`menu_id`
WHERE EXISTS (
    SELECT 1 FROM `nnroad_auth`.`sys_menu` t
    WHERE t.`menu_id` = rm.`menu_id` + @OFFICE_MENU_OFFSET AND t.`app_code` = 'office'
);

-- 可选：超管 role_id=1 一键拿全部 Office 菜单
-- INSERT IGNORE INTO `nnroad_auth`.`sys_role_menu` (`role_id`, `menu_id`)
-- SELECT 1, `menu_id` FROM `nnroad_auth`.`sys_menu` WHERE `app_code` = 'office';

SELECT `app_code`, COUNT(*) AS menu_cnt FROM `sys_menu` GROUP BY `app_code`;
