-- 补拷：角色适用用户类型（Office/Portal 用户管理按 userType 筛角色依赖此表）
-- 在 MySQL 执行，目标库：nnroad_auth

USE `nnroad_auth`;

-- 1) 建表（若 portal 仍有该表，优先 LIKE 拷结构）
CREATE TABLE IF NOT EXISTS `nnroad_auth`.`sys_role_user_type` (
  `role_id` bigint NOT NULL COMMENT '角色ID',
  `user_type` varchar(2) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NOT NULL COMMENT '用户类型',
  PRIMARY KEY (`role_id`, `user_type`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8mb4 COLLATE = utf8mb4_general_ci COMMENT = '角色适用用户类型表' ROW_FORMAT = DYNAMIC;

-- 2) 从 portal 拷数据（表存在且有数据时）
INSERT IGNORE INTO `nnroad_auth`.`sys_role_user_type` (`role_id`, `user_type`)
SELECT `role_id`, `user_type`
FROM `nnroad_portal`.`sys_role_user_type`;

-- 3) 兜底：为 auth 中尚未映射的角色补默认适用类型
--    超管角色 -> 00；其余未映射角色默认给内部员工 01（可再按业务调整）
INSERT IGNORE INTO `nnroad_auth`.`sys_role_user_type` (`role_id`, `user_type`)
SELECT r.role_id, '00'
FROM `nnroad_auth`.`sys_role` r
WHERE r.del_flag = '0'
  AND r.role_key = 'admin'
  AND NOT EXISTS (
    SELECT 1 FROM `nnroad_auth`.`sys_role_user_type` rut WHERE rut.role_id = r.role_id
  );

INSERT IGNORE INTO `nnroad_auth`.`sys_role_user_type` (`role_id`, `user_type`)
SELECT r.role_id, '01'
FROM `nnroad_auth`.`sys_role` r
WHERE r.del_flag = '0'
  AND r.role_key <> 'admin'
  AND NOT EXISTS (
    SELECT 1 FROM `nnroad_auth`.`sys_role_user_type` rut WHERE rut.role_id = r.role_id
  );

-- 校验
SELECT user_type, COUNT(*) AS cnt
FROM `nnroad_auth`.`sys_role_user_type`
GROUP BY user_type
ORDER BY user_type;
