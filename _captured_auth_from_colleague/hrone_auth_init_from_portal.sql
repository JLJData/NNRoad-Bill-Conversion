-- ---------------------------------------------------------------------------
-- 将 Portal 认证数据初始化到 hrone_auth（Office 登录切换前执行）
-- 在 MySQL 执行：source hrone_auth_init_from_portal.sql
-- 执行前请备份：mysqldump -u root -p hrone_portal > hrone_portal_backup.sql
-- ---------------------------------------------------------------------------

CREATE DATABASE IF NOT EXISTS `hrone_auth` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;

-- 以下表从 hrone_portal 复制结构与数据（保留 user_id，portal_user_client 才能关联）
-- 请先执行 sys_menu_app_code.sql 添加 app_code 字段

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_user` LIKE `hrone_portal`.`sys_user`;
INSERT INTO `hrone_auth`.`sys_user` SELECT * FROM `hrone_portal`.`sys_user`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_role` LIKE `hrone_portal`.`sys_role`;
INSERT INTO `hrone_auth`.`sys_role` SELECT * FROM `hrone_portal`.`sys_role`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_menu` LIKE `hrone_portal`.`sys_menu`;
INSERT INTO `hrone_auth`.`sys_menu` SELECT * FROM `hrone_portal`.`sys_menu`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_dept` LIKE `hrone_portal`.`sys_dept`;
INSERT INTO `hrone_auth`.`sys_dept` SELECT * FROM `hrone_portal`.`sys_dept`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_post` LIKE `hrone_portal`.`sys_post`;
INSERT INTO `hrone_auth`.`sys_post` SELECT * FROM `hrone_portal`.`sys_post`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_user_role` LIKE `hrone_portal`.`sys_user_role`;
INSERT INTO `hrone_auth`.`sys_user_role` SELECT * FROM `hrone_portal`.`sys_user_role`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_role_menu` LIKE `hrone_portal`.`sys_role_menu`;
INSERT INTO `hrone_auth`.`sys_role_menu` SELECT * FROM `hrone_portal`.`sys_role_menu`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_role_dept` LIKE `hrone_portal`.`sys_role_dept`;
INSERT INTO `hrone_auth`.`sys_role_dept` SELECT * FROM `hrone_portal`.`sys_role_dept`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_user_post` LIKE `hrone_portal`.`sys_user_post`;
INSERT INTO `hrone_auth`.`sys_user_post` SELECT * FROM `hrone_portal`.`sys_user_post`;

CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_logininfor` LIKE `hrone_portal`.`sys_logininfor`;
INSERT INTO `hrone_auth`.`sys_logininfor` SELECT * FROM `hrone_portal`.`sys_logininfor`;

-- Portal 独有：密码重置（若表存在）
CREATE TABLE IF NOT EXISTS `hrone_auth`.`sys_password_reset` LIKE `hrone_portal`.`sys_password_reset`;
INSERT INTO `hrone_auth`.`sys_password_reset` SELECT * FROM `hrone_portal`.`sys_password_reset`;

-- 校验
SELECT user_type, COUNT(*) AS cnt FROM `hrone_auth`.`sys_user` GROUP BY user_type;
