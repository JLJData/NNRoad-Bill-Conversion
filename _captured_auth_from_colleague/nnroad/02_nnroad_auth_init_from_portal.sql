-- ---------------------------------------------------------------------------
-- 02 创建 nnroad_auth，并从 nnroad_portal 拷贝认证相关表
-- 执行前备份：mysqldump -u root -p nnroad_portal > nnroad_portal_backup.sql
-- 前提：已执行 01_sys_menu_app_code.sql（portal 侧）
-- 注意：保留 user_id，业务表（portal_user_client 等）才能继续关联
-- ---------------------------------------------------------------------------

CREATE DATABASE IF NOT EXISTS `nnroad_auth` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_user` LIKE `nnroad_portal`.`sys_user`;
INSERT INTO `nnroad_auth`.`sys_user` SELECT * FROM `nnroad_portal`.`sys_user`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_role` LIKE `nnroad_portal`.`sys_role`;
INSERT INTO `nnroad_auth`.`sys_role` SELECT * FROM `nnroad_portal`.`sys_role`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_menu` LIKE `nnroad_portal`.`sys_menu`;
INSERT INTO `nnroad_auth`.`sys_menu` SELECT * FROM `nnroad_portal`.`sys_menu`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_dept` LIKE `nnroad_portal`.`sys_dept`;
INSERT INTO `nnroad_auth`.`sys_dept` SELECT * FROM `nnroad_portal`.`sys_dept`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_post` LIKE `nnroad_portal`.`sys_post`;
INSERT INTO `nnroad_auth`.`sys_post` SELECT * FROM `nnroad_portal`.`sys_post`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_user_role` LIKE `nnroad_portal`.`sys_user_role`;
INSERT INTO `nnroad_auth`.`sys_user_role` SELECT * FROM `nnroad_portal`.`sys_user_role`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_role_menu` LIKE `nnroad_portal`.`sys_role_menu`;
INSERT INTO `nnroad_auth`.`sys_role_menu` SELECT * FROM `nnroad_portal`.`sys_role_menu`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_role_dept` LIKE `nnroad_portal`.`sys_role_dept`;
INSERT INTO `nnroad_auth`.`sys_role_dept` SELECT * FROM `nnroad_portal`.`sys_role_dept`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_user_post` LIKE `nnroad_portal`.`sys_user_post`;
INSERT INTO `nnroad_auth`.`sys_user_post` SELECT * FROM `nnroad_portal`.`sys_user_post`;

CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_logininfor` LIKE `nnroad_portal`.`sys_logininfor`;
INSERT INTO `nnroad_auth`.`sys_logininfor` SELECT * FROM `nnroad_portal`.`sys_logininfor`;

-- Portal 密码重置（若表存在）
CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_password_reset` LIKE `nnroad_portal`.`sys_password_reset`;
INSERT INTO `nnroad_auth`.`sys_password_reset` SELECT * FROM `nnroad_portal`.`sys_password_reset`;

-- 角色适用用户类型（用户管理按 userType 筛角色依赖此表；详见 04_nnroad_auth_sys_role_user_type.sql）
CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_role_user_type` LIKE `nnroad_portal`.`sys_role_user_type`;
INSERT IGNORE INTO `nnroad_auth`.`sys_role_user_type` SELECT * FROM `nnroad_portal`.`sys_role_user_type`;

-- 可选：把 Office 独有账号也并进来（仅当 nnroad-office.sys_user 有 portal 没有的账号）
-- INSERT INTO `nnroad_auth`.`sys_user`
-- SELECT o.* FROM `nnroad-office`.`sys_user` o
-- WHERE NOT EXISTS (SELECT 1 FROM `nnroad_auth`.`sys_user` a WHERE a.user_id = o.user_id OR a.user_name = o.user_name);

SELECT user_type, COUNT(*) AS cnt FROM `nnroad_auth`.`sys_user` GROUP BY user_type;
SELECT app_code, COUNT(*) AS cnt FROM `nnroad_auth`.`sys_menu` GROUP BY app_code;
