-- Create syntax for TABLE 'cells'
CREATE TABLE `cells` (
  `id_file` int(11) unsigned NOT NULL,
  `cell` varchar(11) DEFAULT NULL,
  `input` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Create syntax for TABLE 'files'
CREATE TABLE `files` (
  `id` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `nom` varchar(255) DEFAULT NULL,
  `columns` int(11) DEFAULT NULL,
  `rows` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Create syntax for TABLE 'styles'
CREATE TABLE `styles` (
  `id_file` int(11) unsigned NOT NULL,
  `cell` varchar(11) DEFAULT NULL,
  `group` varchar(11) DEFAULT NULL,
  `style` varchar(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;