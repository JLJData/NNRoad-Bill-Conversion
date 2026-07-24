-- ---------------------------------------------------------------------------
-- 01 给菜单加 app_code
-- 在 Portal / Office 源库各执行一次（列已存在会报错，可忽略或先检查）
-- ---------------------------------------------------------------------------

-- Portal
ALTER TABLE `nnroad_portal`.`sys_menu`
    ADD COLUMN `app_code` varchar(32) NOT NULL DEFAULT 'portal'
        COMMENT '应用编码：office=Office工作台，portal=客户门户'
        AFTER `menu_id`;

UPDATE `nnroad_portal`.`sys_menu` SET `app_code` = 'portal' WHERE `app_code` IS NULL OR `app_code` = '';

-- Office（库名：nnroad-office）
ALTER TABLE `nnroad-office`.`sys_menu`
    ADD COLUMN `app_code` varchar(32) NOT NULL DEFAULT 'office'
        COMMENT '应用编码：office=Office工作台，portal=客户门户'
        AFTER `menu_id`;

UPDATE `nnroad-office`.`sys_menu` SET `app_code` = 'office' WHERE `app_code` IS NULL OR `app_code` = '';
