CREATE SCHEMA mutant;

CREATE TABLE `mutant`.`pais` (
  `idPAIS` INT NOT NULL AUTO_INCREMENT,
  `CapitalCity` VARCHAR(150) NULL,
  `ContinentCode` VARCHAR(2) NULL,
  `CountryFlag` VARCHAR(450) NULL,
  `CurrencyISOCode` VARCHAR(3) NULL,
  `ISOCode` VARCHAR(2) NULL,
  `Name` VARCHAR(120) NULL,
  `PhoneCode` INT NULL,
  PRIMARY KEY (`idPAIS`))
COMMENT = 'Dados sobre o PAIS capturado da API http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries';


CREATE TABLE `mutant`.`idioma` (
  `idIdioma` INT NOT NULL AUTO_INCREMENT,
  `IsoCode` VARCHAR(3) NULL,
  `Name` VARCHAR(120) NULL,
  PRIMARY KEY (`idIdioma`))
COMMENT = 'Dados sobre o IDIOMA capturado da API http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries';

CREATE TABLE `mutant`.`idiomas_do_pais` (
  `idIdiomas_do_pais` INT NOT NULL AUTO_INCREMENT,
  `idPais` INT NULL,
  `idIdioma` INT NULL,
  PRIMARY KEY (`idIdiomas_do_pais`),
  INDEX `idPais_idx` (`idPais` ASC) VISIBLE,
  INDEX `IdIdioma_idx` (`idIdioma` ASC) VISIBLE,
  CONSTRAINT `idPais`
    FOREIGN KEY (`idPais`)
    REFERENCES `mutant`.`pais` (`idPAIS`)
    ON DELETE NO ACTION
    ON UPDATE NO ACTION,
  CONSTRAINT `IdIdioma`
    FOREIGN KEY (`idIdioma`)
    REFERENCES `mutant`.`idioma` (`idIdioma`)
    ON DELETE NO ACTION
    ON UPDATE NO ACTION)
COMMENT = 'Dados sobre o IDIOMA FALADO no PAIS capturado da API http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries';
