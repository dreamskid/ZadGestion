-- phpMyAdmin SQL Dump
-- version 4.7.0
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Nov 25, 2017 at 05:08 PM
-- Server version: 10.1.25-MariaDB
-- PHP Version: 5.6.31

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `zadgestion`
--

-- --------------------------------------------------------

--
-- Table structure for table `clients`
--

CREATE TABLE `clients` (
  `address` varchar(256) NOT NULL,
  `city` varchar(256) NOT NULL,
  `country` varchar(256) NOT NULL,
  `corporate_name` varchar(256) NOT NULL,
  `corporate_number` varchar(256) NOT NULL,
  `date_creation` varchar(256) NOT NULL,
  `email` varchar(256) NOT NULL,
  `id` varchar(256) NOT NULL,
  `phone` varchar(256) NOT NULL,
  `state` varchar(256) NOT NULL,
  `zipcode` varchar(256) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `hostsandhostesses`
--

CREATE TABLE `hostsandhostesses` (
  `address` varchar(256) NOT NULL,
  `birth_city` varchar(256) NOT NULL,
  `birth_date` varchar(256) NOT NULL,
  `cellphone` varchar(256) NOT NULL,
  `city` varchar(256) NOT NULL,
  `country` varchar(256) NOT NULL,
  `date_creation` varchar(256) NOT NULL,
  `email` varchar(256) NOT NULL,
  `firstname` varchar(256) NOT NULL,
  `has_car` tinyint(1) NOT NULL,
  `has_driver_licence` tinyint(1) NOT NULL,
  `id` varchar(256) NOT NULL,
  `id_paycheck` varchar(256) NOT NULL,
  `language_english` tinyint(1) NOT NULL,
  `language_german` tinyint(1) NOT NULL,
  `language_italian` tinyint(1) NOT NULL,
  `language_others` varchar(256) NOT NULL,
  `language_spanish` tinyint(1) NOT NULL,
  `lastname` varchar(256) NOT NULL,
  `profile_event` tinyint(1) NOT NULL,
  `profile_permanent` tinyint(1) NOT NULL,
  `profile_street` tinyint(1) NOT NULL,
  `sex` varchar(256) NOT NULL,
  `size` varchar(256) NOT NULL,
  `size_pants` varchar(2) NOT NULL,
  `size_shirt` varchar(2) NOT NULL,
  `size_shoes` varchar(2) NOT NULL,
  `social_number` varchar(256) NOT NULL,
  `state` varchar(256) NOT NULL,
  `zipcode` varchar(256) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `hostsandhostesses`
--

INSERT INTO `hostsandhostesses` (`address`, `birth_city`, `birth_date`, `cellphone`, `city`, `country`, `date_creation`, `email`, `firstname`, `has_car`, `has_driver_licence`, `id`, `id_paycheck`, `language_english`, `language_german`, `language_italian`, `language_others`, `language_spanish`, `lastname`, `profile_event`, `profile_permanent`, `profile_street`, `sex`, `size`, `size_pants`, `size_shirt`, `size_shoes`, `social_number`, `state`, `zipcode`) VALUES
('15 bis rue Roussy', 'Limoges', '3 Janvier 1985', '0760071696', 'Lyon', 'France', '2017-11-11 00:00:00 ', 'juliette.sailland@gmail.com', 'Juliette', 0, 0, '0_JS_69004', '3535252342352353235', 1, 1, 0, '', 0, 'Sailland', 1, 0, 1, 'F', '165', '34', 'S', '37', '285011702901831', '', '69004'),
('12 rue de Saint Germain', 'Lyon', '5 Février 1982', '0699304795', 'L\'Isle-d\'Abeau', 'France', '2017-11-12 00:00:00 ', 'yohann.tschudi@gmail.com', 'Yohann', 0, 0, '1_YT_69004', '353543523523523523', 1, 0, 0, 'Russe', 1, 'Tschudi', 1, 1, 0, 'M', '185', '42', 'L', '42', '182026902902791', '', '38080'),
('12 rue Maurice Chevalier', 'Saint Brieuc', '2 Août 1989', '0660777459', 'Lamballe', 'France', '2017-11-13 00:00:00 ', 'pierre.lesnard@zadigandco.com', 'Pierre', 0, 1, '2_PL_22400', '34235325353253532', 1, 0, 0, '', 0, 'Lesnard', 1, 1, 1, 'M', '183', '38', 'M', '44', '124341432312332', '', '22400');

-- --------------------------------------------------------

--
-- Table structure for table `settings`
--

CREATE TABLE `settings` (
  `id` int(11) NOT NULL,
  `photos_repository` varchar(256) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `settings`
--

INSERT INTO `settings` (`id`, `photos_repository`) VALUES
(0, 'D:\\Yohann\\Zadig & Co\\Photos');
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
