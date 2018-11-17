-- phpMyAdmin SQL Dump
-- version 4.7.4
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: May 19, 2018 at 10:08 AM
-- Server version: 10.1.26-MariaDB
-- PHP Version: 7.1.9

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
  `archived` int(1) NOT NULL,
  `city` varchar(256) NOT NULL,
  `country` varchar(256) NOT NULL,
  `corporate_name` varchar(256) NOT NULL,
  `corporate_number` varchar(256) NOT NULL,
  `date_creation` varchar(256) NOT NULL,
  `id` varchar(256) NOT NULL,
  `phone` varchar(256) NOT NULL,
  `state` varchar(256) NOT NULL,
  `vat_number` varchar(256) NOT NULL,
  `zipcode` varchar(256) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `clients`
--

INSERT INTO `clients` (`address`, `archived`, `city`, `country`, `corporate_name`, `corporate_number`, `date_creation`, `id`, `phone`, `state`, `vat_number`, `zipcode`) VALUES
('Immeuble Le quartz', 0, 'Villeurbanne', 'France', 'Yole Développement', '16146513', '2017-11-26 00:00:00 ', '2_Yo_16_69100', '0485128613', '', '2737837373543', '69100'),
('12 rue de Meylan', 1, 'Nice', 'France', 'Fantôme Inc.', '1641616516416', '2017-12-11 00:00:00 ', '2_Fa_16_06000', '0225654132', '', '', '06000'),
('115 rue du genou', 1, 'Paris', 'France', 'Test', '164516513', '2018-01-05 00:00:00 ', '3_Te_16_75010', '0485659812', '', '', '75010');

-- --------------------------------------------------------

--
-- Table structure for table `hostsandhostesses`
--

CREATE TABLE `hostsandhostesses` (
  `address` varchar(256) NOT NULL,
  `archived` int(1) NOT NULL,
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

INSERT INTO `hostsandhostesses` (`address`, `archived`, `birth_city`, `birth_date`, `cellphone`, `city`, `country`, `date_creation`, `email`, `firstname`, `has_car`, `has_driver_licence`, `id`, `id_paycheck`, `language_english`, `language_german`, `language_italian`, `language_others`, `language_spanish`, `lastname`, `profile_event`, `profile_permanent`, `profile_street`, `sex`, `size`, `size_pants`, `size_shirt`, `size_shoes`, `social_number`, `state`, `zipcode`) VALUES
('15 bis rue Roussy', 0, 'Limoges', '3 Janvier 1985', '0760071696', 'Lyon', 'France', '2017-11-11 00:00:00 ', 'juliette.sailland@gmail.com', 'Juliette', 0, 0, '0_JS_69004', '3535252342352353235', 1, 0, 0, '', 0, 'Sailland', 1, 0, 1, 'F', '165', '34', 'S', '37', '285011702901831', '', '69004'),
('12 rue de Saint Germain', 0, 'Lyon', '5 Février 1982', '0699304795', 'L\'Isle-d\'Abeau', 'France', '2017-11-12 00:00:00 ', 'yohann.tschudi@gmail.com', 'Yohann', 0, 0, '1_YT_69004', '353543523523523523', 1, 0, 0, 'Russe', 1, 'Tschudi', 1, 1, 0, 'M', '185', '42', 'L', '42', '182026902902791', '', '38080'),
('12 rue Maurice Chevalier', 1, 'Saint Brieuc', '2 Août 1989', '0660777459', 'Lamballe', 'France', '2017-11-13 00:00:00 ', 'pierre.lesnard@zadigandco.com', 'Pierre', 0, 1, '2_PL_22400', '34235325353253532', 1, 0, 0, '', 0, 'Lesnard', 1, 1, 1, 'M', '183', '38', 'M', '44', '124341432312332', '', '22400'),
('12 rue du Cuire', 1, 'Bron', '7 Février 1995', '0684848484', 'Lyon', 'France', '2018-01-05 00:00:00 ', 'michel.test@gmail.com', 'Michel', 1, 1, '3_MT_69004', '211351351351', 1, 0, 0, 'Russe, Chinois', 0, 'Test', 0, 0, 1, 'M', '180', '42', 'L', '42', '1854545651265', '', '69004');

-- --------------------------------------------------------

--
-- Table structure for table `missions`
--

CREATE TABLE `missions` (
  `address` varchar(256) NOT NULL,
  `archived` int(1) NOT NULL,
  `city` varchar(256) NOT NULL,
  `client_name` varchar(256) NOT NULL,
  `country` varchar(256) NOT NULL,
  `date_creation` varchar(256) NOT NULL,
  `description` varchar(256) NOT NULL,
  `end_date` varchar(256) NOT NULL,
  `id` varchar(256) NOT NULL,
  `id_list_shifts` varchar(256) NOT NULL,
  `start_date` varchar(256) NOT NULL,
  `state` varchar(256) NOT NULL,
  `zipcode` varchar(256) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `missions`
--

INSERT INTO `missions` (`address`, `archived`, `city`, `client_name`, `country`, `date_creation`, `description`, `end_date`, `id`, `id_list_shifts`, `start_date`, `state`, `zipcode`) VALUES
('Immeuble Le Quartz', 0, 'Villeurbanne', 'Yole Développement', 'France', '2018-01-05 00:00:00 ', 'Démonstration de nouveaux produits en magasin devant les clients et leurs familles', '04/02/2018 00:00:00', '0_Y_Villeurbanne', '', '01/02/2018 00:00:00', '', '69100'),
('15 rue de la République', 1, 'Paris', 'Fantôme Inc.', 'France', '2018-01-05 00:00:00 ', 'Distribution flyers', '01/03/2018 00:00:00', '1_F_Paris', '', '28/02/2018 00:00:00', '', '75006'),
('Ullis', 0, 'Grenoble', 'Zadig&Co', 'France', '2018-01-05 00:00:00 ', 'Recherche données et developpement logiciel de prestations', '09/01/2018 00:00:00', '2_Y_Grenoble', '', '08/01/2018 00:00:00', '', '38000'),
('Eurexpo', 1, 'Caluire-et-Cuire', 'Zadig&Co', 'Distribution Flyers', '2018-01-05 00:00:00 ', 'France', '20/01/2018 00:00:00', '3_Z_Caluire-et-Cuire', '', '19/01/2018 00:00:00', '', '69300'),
('Immeuble Le Quartz', 0, 'Villeurbanne', 'Yole Développement', 'France', '2018-01-05 00:00:00 ', 'Démonstration de nouveaux produits en magasin devant les clients et leurs familles', '04/02/2018 00:00:00', '4_Y_Villeurbanne', '', '01/02/2018 00:00:00', '', '69100'),
('Ullis', 0, 'Grenoble', 'Zadig&Co', 'France', '2018-01-05 00:00:00 ', 'Recherche données et developpement logiciel de prestations', '09/01/2018 00:00:00', '5_Z_Grenoble', '5_Z_Grenoble_10/05/2018_08:00_14:00_0_JS_69004', '08/01/2018 00:00:00', '', '38000'),
('ffq', 0, 'Villeurbanne', 'Yole Développement', 'fwewefwe', '2018-05-08 00:00:00 ', 'France', '12/05/2018 00:00:00', '7_Y_Villeurbanne', '', '11/05/2018 00:00:00', '', '69100');

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
(0, 'C:\\Users\\Yohann\\Desktop\\Softwares\\Zadig & Co\\Photos');

-- --------------------------------------------------------

--
-- Table structure for table `shifts`
--

CREATE TABLE `shifts` (
  `date` varchar(256) NOT NULL,
  `end_time` varchar(256) NOT NULL,
  `hourly_rate` varchar(256) NOT NULL,
  `id` varchar(256) NOT NULL,
  `id_hostorhostess` varchar(256) NOT NULL,
  `id_mission` varchar(256) NOT NULL,
  `start_time` varchar(256) NOT NULL,
  `suit` tinyint(1) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `shifts`
--

INSERT INTO `shifts` (`date`, `end_time`, `hourly_rate`, `id`, `id_hostorhostess`, `id_mission`, `start_time`, `suit`) VALUES
('08/05/2018', '17:00', '10', '5_Z_Grenoble_08/05/2018_08:00_17:00_0_JS_69004', '0_JS_69004', '5_Z_Grenoble', '08:00', 1),
('10/05/2018', '14:00', '10', '5_Z_Grenoble_10/05/2018_08:00_14:00_0_JS_69004', '0_JS_69004', '5_Z_Grenoble', '08:00', 1);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
