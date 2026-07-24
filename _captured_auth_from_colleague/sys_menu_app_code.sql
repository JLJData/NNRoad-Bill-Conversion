-- sys_menu ���� app_code������ Office / Portal �˵���Ȩ��
-- �� hrone_auth����ǰ�� sys_menu �Ŀ⣩ִ��

ALTER TABLE `sys_menu`
    ADD COLUMN `app_code` varchar(32) NOT NULL DEFAULT 'portal'
        COMMENT '应用编码：office=Office工作台，portal=客户门户'
        AFTER `menu_id`;

-- 默认 Portal 应用编码为 portal
-- RuoYi 数据初始化时 Office 应用编码为：
-- UPDATE sys_menu SET app_code = 'office' WHERE menu_id IN (...);

-- 校验应用编码分布
SELECT app_code, COUNT(*) AS cnt FROM sys_menu GROUP BY app_code;
