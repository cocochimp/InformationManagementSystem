/*
 Navicat Premium Data Transfer

 Source Server         : localhost_3306
 Source Server Type    : MySQL
 Source Server Version : 50635
 Source Host           : localhost:3306
 Source Schema         : ocean_student

 Target Server Type    : MySQL
 Target Server Version : 50635
 File Encoding         : 65001

 Date: 08/08/2023 18:54:39
*/

SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------
-- Table structure for all_student
-- ----------------------------
DROP TABLE IF EXISTS `all_student`;
CREATE TABLE `all_student`  (
  `id` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `name` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `startYear` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `profession` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `klass` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `number` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of all_student
-- ----------------------------
INSERT INTO `all_student` VALUES ('201711621428', '吴伟盛', '2017', '计算机科学与技术', '计科1174', '28');
INSERT INTO `all_student` VALUES ('201711701130', '谢培烽', '2017', '软件工程', '软工1171', '30');
INSERT INTO `all_student` VALUES ('201711701416', '李耀璋', '2017', '软件工程', '软工1174', '16');
INSERT INTO `all_student` VALUES ('201811671102', '陈  干', '2018', '信息系统与信息管理', '信管1181', '02');
INSERT INTO `all_student` VALUES ('201811701111', '李慧敏', '2018', '软件工程', '软工1181', '11');
INSERT INTO `all_student` VALUES ('201911621208', '丁  滢', '2019', '计算机科学与技术', '计科1192', '08');
INSERT INTO `all_student` VALUES ('201911621218', '林树富', '2019', '计算机科学与技术', '计科1192', '18');
INSERT INTO `all_student` VALUES ('201911701105', '陈松源', '2019', '软件工程', '软工1191', '05');
INSERT INTO `all_student` VALUES ('201911701120', '温嘉慧', '2019', '软件工程', '软工1191', '20');
INSERT INTO `all_student` VALUES ('201911701132', '张劲婷', '2019', '软件工程', '软工1191', '32');
INSERT INTO `all_student` VALUES ('201911921706', '郭嘉麒', '2019', '信息与计算科学', '信计1197', '06');

-- ----------------------------
-- Table structure for books
-- ----------------------------
DROP TABLE IF EXISTS `books`;
CREATE TABLE `books`  (
  `bookID` int(10) NOT NULL AUTO_INCREMENT COMMENT '书id',
  `bookName` varchar(100) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL COMMENT '书名',
  `bookCounts` int(11) NOT NULL COMMENT '数量',
  `detail` varchar(200) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL COMMENT '描述',
  INDEX `bookID`(`bookID`) USING BTREE
) ENGINE = InnoDB AUTO_INCREMENT = 4 CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of books
-- ----------------------------
INSERT INTO `books` VALUES (1, 'MySQL', 10, '从删库到跑路');
INSERT INTO `books` VALUES (2, 'Linux', 5, '从进门到进牢');
INSERT INTO `books` VALUES (3, 'Spring', 2, '从入门到入土');

-- ----------------------------
-- Table structure for snakescore
-- ----------------------------
DROP TABLE IF EXISTS `snakescore`;
CREATE TABLE `snakescore`  (
  `No` int(5) NOT NULL,
  `Score` int(100) NOT NULL,
  PRIMARY KEY (`No`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of snakescore
-- ----------------------------
INSERT INTO `snakescore` VALUES (1, 30);
INSERT INTO `snakescore` VALUES (2, 20);
INSERT INTO `snakescore` VALUES (3, 10);

-- ----------------------------
-- Table structure for users
-- ----------------------------
DROP TABLE IF EXISTS `users`;
CREATE TABLE `users`  (
  `id` int(11) NOT NULL,
  `name` varchar(40) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `password` varchar(40) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `email` varchar(60) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `birthday` date NULL DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Compact;

-- ----------------------------
-- Records of users
-- ----------------------------
INSERT INTO `users` VALUES (1, '张三', '123456', 'zs@qq.com', '2000-01-01');
INSERT INTO `users` VALUES (2, '李四', '123456', 'ls@qq.com', '2000-01-01');
INSERT INTO `users` VALUES (3, '王五', '123456', 'ww@qq.com', '2000-01-01');
INSERT INTO `users` VALUES (4, '狂神说Java', '123456', '24736743@qq.com', '2021-12-05');

SET FOREIGN_KEY_CHECKS = 1;
