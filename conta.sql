# MySQL-Front Dump 2.5
#
# Host: localhost   Database: Conta2
# --------------------------------------------------------
# Server version 3.23.53-max-nt


#
# Table structure for table 'agentes'
#

CREATE TABLE agentes (
  Codigo smallint(5) unsigned NOT NULL default '0',
  Nombre varchar(30) NOT NULL default '',
  PRIMARY KEY  (Codigo)
) TYPE=MyISAM;



#
# Table structure for table 'cabapu'
#

CREATE TABLE cabapu (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien mediumint(1) unsigned NOT NULL default '0',
  bloqactu tinyint(1) NOT NULL default '0',
  numaspre smallint(1) default NULL,
  obsdiari text,
  PRIMARY KEY  (numdiari,fechaent,numasien),
  KEY cl_asientospredefinidos (numaspre),
  FOREIGN KEY (numdiari) REFERENCES conta2.tiposdiario (numdiari),
  FOREIGN KEY (numaspre) REFERENCES conta2.cabasipre (numaspre) ON DELETE SET NULL
) TYPE=InnoDB;



#
# Table structure for table 'cabapue'
#

CREATE TABLE cabapue (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien smallint(1) unsigned NOT NULL default '0',
  bloqactu tinyint(1) NOT NULL default '0',
  numaspre smallint(1) default NULL,
  obsdiari text,
  PRIMARY KEY  (numdiari,fechaent,numasien)
) TYPE=MyISAM;



#
# Table structure for table 'cabasipre'
#

CREATE TABLE cabasipre (
  numaspre smallint(1) NOT NULL default '0',
  nomaspre char(40) NOT NULL default '',
  PRIMARY KEY  (numaspre)
) TYPE=InnoDB;



#
# Table structure for table 'cabccost'
#

CREATE TABLE cabccost (
  codccost char(4) NOT NULL default '0',
  nomccost char(30) NOT NULL default '0',
  idsubcos tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (codccost)
) TYPE=InnoDB;



#
# Table structure for table 'cabfact'
#

CREATE TABLE cabfact (
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  fecfaccl date NOT NULL default '0000-00-00',
  codmacta char(10) NOT NULL default '',
  anofaccl smallint(6) NOT NULL default '0',
  confaccl char(15) default NULL,
  ba1faccl decimal(12,2) NOT NULL default '0.00',
  ba2faccl decimal(12,2) default NULL,
  ba3faccl decimal(12,2) default NULL,
  pi1faccl decimal(6,2) default NULL,
  pi2faccl decimal(6,2) default NULL,
  pi3faccl decimal(6,2) default NULL,
  pr1faccl decimal(6,2) default NULL,
  pr2faccl decimal(6,2) default NULL,
  pr3faccl decimal(6,2) default NULL,
  ti1faccl decimal(12,2) default NULL,
  ti2faccl decimal(12,2) default NULL,
  ti3faccl decimal(12,2) default NULL,
  tr1faccl decimal(12,2) default NULL,
  tr2faccl decimal(12,2) default NULL,
  tr3faccl decimal(12,2) default NULL,
  totfaccl decimal(14,2) default NULL,
  tp1faccl tinyint(1) unsigned NOT NULL default '0',
  tp2faccl tinyint(3) unsigned default NULL,
  tp3faccl tinyint(3) unsigned default NULL,
  intracom tinyint(3) unsigned NOT NULL default '0',
  retfaccl decimal(6,2) default NULL,
  trefaccl decimal(12,2) default NULL,
  cuereten char(10) default NULL,
  numdiari smallint(1) unsigned default NULL,
  fechaent date default NULL,
  numasien mediumint(1) unsigned default NULL,
  fecliqcl date NOT NULL default '0000-00-00',
  PRIMARY KEY  (numserie,codfaccl,anofaccl),
  KEY cl_TipoIVA1 (tp1faccl),
  KEY cl_TipoIVA2 (tp2faccl),
  KEY cl_TipoIVA3 (tp3faccl),
  KEY cl_facCodmacta (codmacta),
  KEY cl_faccuereten (cuereten),
  KEY cl_Contadores (numserie),
  FOREIGN KEY (numserie) REFERENCES conta2.contadores (tiporegi),
  FOREIGN KEY (tp1faccl) REFERENCES conta2.tiposiva (codigiva),
  FOREIGN KEY (tp2faccl) REFERENCES conta2.tiposiva (codigiva),
  FOREIGN KEY (tp3faccl) REFERENCES conta2.tiposiva (codigiva),
  FOREIGN KEY (cuereten) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'cabfact1'
#

CREATE TABLE cabfact1 (
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  fecfaccl date NOT NULL default '0000-00-00',
  codmacta char(10) NOT NULL default '',
  anofaccl smallint(6) NOT NULL default '0',
  confaccl char(15) default NULL,
  ba1faccl decimal(12,2) NOT NULL default '0.00',
  ba2faccl decimal(12,2) default NULL,
  ba3faccl decimal(12,2) default NULL,
  pi1faccl decimal(6,2) default NULL,
  pi2faccl decimal(6,2) default NULL,
  pi3faccl decimal(6,2) default NULL,
  pr1faccl decimal(6,2) default NULL,
  pr2faccl decimal(6,2) default NULL,
  pr3faccl decimal(6,2) default NULL,
  ti1faccl decimal(12,2) default NULL,
  ti2faccl decimal(12,2) default NULL,
  ti3faccl decimal(12,2) default NULL,
  tr1faccl decimal(12,2) default NULL,
  tr2faccl decimal(12,2) default NULL,
  tr3faccl decimal(12,2) default NULL,
  totfaccl decimal(14,2) default NULL,
  tp1faccl tinyint(1) unsigned NOT NULL default '0',
  tp2faccl tinyint(3) unsigned default NULL,
  tp3faccl tinyint(3) unsigned default NULL,
  intracom tinyint(3) unsigned NOT NULL default '0',
  retfaccl decimal(6,2) default NULL,
  trefaccl decimal(12,2) default NULL,
  cuereten char(10) default NULL,
  numdiari smallint(1) unsigned default NULL,
  fechaent date default NULL,
  numasien mediumint(1) unsigned default NULL,
  fecliqcl date NOT NULL default '0000-00-00',
  PRIMARY KEY  (numserie,codfaccl,anofaccl)
) TYPE=MyISAM;



#
# Table structure for table 'cabfacte'
#

CREATE TABLE cabfacte (
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  fecfaccl date NOT NULL default '0000-00-00',
  codmacta char(10) NOT NULL default '',
  anofaccl smallint(6) NOT NULL default '0',
  confaccl char(15) default NULL,
  ba1faccl decimal(12,2) NOT NULL default '0.00',
  ba2faccl decimal(12,2) default NULL,
  ba3faccl decimal(12,2) default NULL,
  pi1faccl decimal(6,2) default NULL,
  pi2faccl decimal(6,2) default NULL,
  pi3faccl decimal(6,2) default NULL,
  pr1faccl decimal(6,2) default NULL,
  pr2faccl decimal(6,2) default NULL,
  pr3faccl decimal(6,2) default NULL,
  ti1faccl decimal(12,2) default NULL,
  ti2faccl decimal(12,2) default NULL,
  ti3faccl decimal(12,2) default NULL,
  tr1faccl decimal(12,2) default NULL,
  tr2faccl decimal(12,2) default NULL,
  tr3faccl decimal(12,2) default NULL,
  totfaccl decimal(14,2) default NULL,
  tp1faccl tinyint(1) unsigned NOT NULL default '0',
  tp2faccl tinyint(3) unsigned default NULL,
  tp3faccl tinyint(3) unsigned default NULL,
  intracom tinyint(3) unsigned NOT NULL default '0',
  retfaccl decimal(6,2) default NULL,
  trefaccl decimal(12,2) default NULL,
  cuereten char(10) default NULL,
  numdiari smallint(1) unsigned default NULL,
  fechaent date default NULL,
  numasien mediumint(1) unsigned default NULL,
  fecliqcl date NOT NULL default '0000-00-00',
  PRIMARY KEY  (numserie,codfaccl,anofaccl)
) TYPE=MyISAM;



#
# Table structure for table 'cabfactprov'
#

CREATE TABLE cabfactprov (
  numregis int(11) NOT NULL default '0',
  fecfacpr date NOT NULL default '0000-00-00',
  anofacpr smallint(6) NOT NULL default '0',
  fecrecpr date NOT NULL default '0000-00-00',
  numfacpr char(10) NOT NULL default '',
  codmacta char(10) NOT NULL default '',
  confacpr char(15) default NULL,
  ba1facpr decimal(12,2) NOT NULL default '0.00',
  ba2facpr decimal(12,2) default NULL,
  ba3facpr decimal(12,2) default NULL,
  pi1facpr decimal(6,2) default NULL,
  pi2facpr decimal(6,2) default NULL,
  pi3facpr decimal(6,2) default NULL,
  pr1facpr decimal(6,2) default NULL,
  pr2facpr decimal(6,2) default NULL,
  pr3facpr decimal(6,2) default NULL,
  ti1facpr decimal(12,2) default NULL,
  ti2facpr decimal(12,2) default NULL,
  ti3facpr decimal(12,2) default NULL,
  tr1facpr decimal(12,2) default NULL,
  tr2facpr decimal(12,2) default NULL,
  tr3facpr decimal(12,2) default NULL,
  totfacpr decimal(14,2) default NULL,
  tp1facpr tinyint(1) unsigned NOT NULL default '0',
  tp2facpr tinyint(3) unsigned default NULL,
  tp3facpr tinyint(3) unsigned default NULL,
  extranje tinyint(3) unsigned NOT NULL default '0',
  retfacpr decimal(6,2) default NULL,
  trefacpr decimal(12,2) default NULL,
  cuereten char(10) default NULL,
  numdiari smallint(1) unsigned default NULL,
  fechaent date default NULL,
  numasien mediumint(1) unsigned default NULL,
  fecliqpr date NOT NULL default '0000-00-00',
  nodeducible tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (numregis,anofacpr),
  KEY pr_TipoIVA1 (tp1facpr),
  KEY pr_TipoIVA2 (tp2facpr),
  KEY pr_TipoIVA3 (tp3facpr),
  KEY pr_facCodmacta (codmacta),
  KEY pr_faccuereten (cuereten),
  FOREIGN KEY (tp1facpr) REFERENCES conta2.tiposiva (codigiva),
  FOREIGN KEY (tp2facpr) REFERENCES conta2.tiposiva (codigiva),
  FOREIGN KEY (tp3facpr) REFERENCES conta2.tiposiva (codigiva),
  FOREIGN KEY (cuereten) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'cabfactprov1'
#

CREATE TABLE cabfactprov1 (
  numregis int(11) NOT NULL default '0',
  fecfacpr date NOT NULL default '0000-00-00',
  anofacpr smallint(6) NOT NULL default '0',
  fecrecpr date NOT NULL default '0000-00-00',
  numfacpr char(10) NOT NULL default '',
  codmacta char(10) NOT NULL default '',
  confacpr char(15) default NULL,
  ba1facpr decimal(12,2) NOT NULL default '0.00',
  ba2facpr decimal(12,2) default NULL,
  ba3facpr decimal(12,2) default NULL,
  pi1facpr decimal(6,2) default NULL,
  pi2facpr decimal(6,2) default NULL,
  pi3facpr decimal(6,2) default NULL,
  pr1facpr decimal(6,2) default NULL,
  pr2facpr decimal(6,2) default NULL,
  pr3facpr decimal(6,2) default NULL,
  ti1facpr decimal(12,2) default NULL,
  ti2facpr decimal(12,2) default NULL,
  ti3facpr decimal(12,2) default NULL,
  tr1facpr decimal(12,2) default NULL,
  tr2facpr decimal(12,2) default NULL,
  tr3facpr decimal(12,2) default NULL,
  totfacpr decimal(14,2) default NULL,
  tp1facpr tinyint(1) unsigned NOT NULL default '0',
  tp2facpr tinyint(3) unsigned default NULL,
  tp3facpr tinyint(3) unsigned default NULL,
  extranje tinyint(3) unsigned NOT NULL default '0',
  retfacpr decimal(6,2) default NULL,
  trefacpr decimal(12,2) default NULL,
  cuereten char(10) default NULL,
  numdiari smallint(1) unsigned default NULL,
  fechaent date default NULL,
  numasien mediumint(1) unsigned default NULL,
  fecliqpr date NOT NULL default '0000-00-00',
  nodeducible tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (numregis,anofacpr)
) TYPE=MyISAM;



#
# Table structure for table 'cabfactprove'
#

CREATE TABLE cabfactprove (
  numregis int(11) NOT NULL default '0',
  fecfacpr date NOT NULL default '0000-00-00',
  anofacpr smallint(6) NOT NULL default '0',
  fecrecpr date NOT NULL default '0000-00-00',
  numfacpr char(10) NOT NULL default '',
  codmacta char(10) NOT NULL default '',
  confacpr char(15) default NULL,
  ba1facpr decimal(12,2) NOT NULL default '0.00',
  ba2facpr decimal(12,2) default NULL,
  ba3facpr decimal(12,2) default NULL,
  pi1facpr decimal(6,2) default NULL,
  pi2facpr decimal(6,2) default NULL,
  pi3facpr decimal(6,2) default NULL,
  pr1facpr decimal(6,2) default NULL,
  pr2facpr decimal(6,2) default NULL,
  pr3facpr decimal(6,2) default NULL,
  ti1facpr decimal(12,2) default NULL,
  ti2facpr decimal(12,2) default NULL,
  ti3facpr decimal(12,2) default NULL,
  tr1facpr decimal(12,2) default NULL,
  tr2facpr decimal(12,2) default NULL,
  tr3facpr decimal(12,2) default NULL,
  totfacpr decimal(14,2) default NULL,
  tp1facpr tinyint(1) unsigned NOT NULL default '0',
  tp2facpr tinyint(3) unsigned default NULL,
  tp3facpr tinyint(3) unsigned default NULL,
  extranje tinyint(3) unsigned NOT NULL default '0',
  retfacpr decimal(6,2) default NULL,
  trefacpr decimal(12,2) default NULL,
  cuereten char(10) default NULL,
  numdiari smallint(1) unsigned default NULL,
  fechaent date default NULL,
  numasien mediumint(1) unsigned default NULL,
  fecliqpr date NOT NULL default '0000-00-00',
  nodeducible tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (numregis,anofacpr)
) TYPE=MyISAM;



#
# Table structure for table 'conceptos'
#

CREATE TABLE conceptos (
  codconce smallint(1) NOT NULL default '0',
  nomconce char(30) NOT NULL default '0',
  tipoconce tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (codconce)
) TYPE=InnoDB;



#
# Table structure for table 'contadores'
#

CREATE TABLE contadores (
  tiporegi char(1) NOT NULL default '',
  nomregis char(30) NOT NULL default '',
  contado1 mediumint(9) default NULL,
  contado2 mediumint(9) default NULL,
  PRIMARY KEY  (tiporegi)
) TYPE=InnoDB;



#
# Table structure for table 'contadoreshco'
#

CREATE TABLE contadoreshco (
  anoregis smallint(1) unsigned NOT NULL default '0',
  tiporegi char(1) NOT NULL default '',
  nomregis char(30) NOT NULL default '',
  contado1 mediumint(9) default NULL,
  contado2 mediumint(9) default NULL,
  PRIMARY KEY  (tiporegi,anoregis)
) TYPE=MyISAM;



#
# Table structure for table 'ctaagrupadas'
#

CREATE TABLE ctaagrupadas (
  codmacta char(10) NOT NULL default '',
  PRIMARY KEY  (codmacta),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'ctabancaria'
#

CREATE TABLE ctabancaria (
  codmacta varchar(10) NOT NULL default '0',
  entidad smallint(1) unsigned NOT NULL default '0',
  oficina smallint(1) unsigned NOT NULL default '0',
  control char(2) default NULL,
  ctabanco varchar(10) NOT NULL default '0',
  descripcion varchar(40) default NULL,
  sufijoem char(3) default NULL,
  PRIMARY KEY  (codmacta)
) TYPE=MyISAM;



#
# Table structure for table 'ctaexclusion'
#

CREATE TABLE ctaexclusion (
  codmacta char(10) NOT NULL default '',
  PRIMARY KEY  (codmacta),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'cuentas'
#

CREATE TABLE cuentas (
  codmacta varchar(10) NOT NULL default '0',
  nommacta varchar(30) NOT NULL default '0',
  apudirec char(1) NOT NULL default '0',
  model347 tinyint(1) NOT NULL default '0',
  razosoci varchar(30) default '',
  dirdatos varchar(30) default '',
  codposta varchar(6) default '',
  despobla varchar(30) default '',
  desprovi varchar(30) default NULL,
  nifdatos varchar(15) default '',
  maidatos varchar(50) default '',
  webdatos varchar(50) default '',
  obsdatos text,
  pais varchar(15) default NULL,
  PRIMARY KEY  (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'departamentos'
#

CREATE TABLE departamentos (
  codmacta varchar(10) NOT NULL default '',
  Dpto tinyint(3) unsigned NOT NULL default '0',
  Descripcion varchar(30) NOT NULL default '0',
  PRIMARY KEY  (codmacta,Dpto)
) TYPE=MyISAM;



#
# Table structure for table 'empresa'
#

CREATE TABLE empresa (
  codempre char(8) NOT NULL default '0',
  nomempre char(40) NOT NULL default '',
  nomresum char(15) default NULL,
  numnivel tinyint(1) unsigned NOT NULL default '0',
  numdigi1 tinyint(1) unsigned default '0',
  numdigi2 tinyint(1) unsigned NOT NULL default '0',
  numdigi3 tinyint(1) unsigned default '0',
  numdigi4 tinyint(1) unsigned default '0',
  numdigi5 tinyint(1) unsigned default '0',
  numdigi6 tinyint(1) unsigned default '0',
  numdigi7 tinyint(1) unsigned default '0',
  numdigi8 tinyint(1) unsigned default '0',
  numdigi9 tinyint(1) unsigned default '0',
  numdigi10 tinyint(1) unsigned default '0',
  PRIMARY KEY  (codempre)
) TYPE=InnoDB;



#
# Table structure for table 'empresa2'
#

CREATE TABLE empresa2 (
  codigo tinyint(4) NOT NULL default '0',
  apoderado char(100) default NULL,
  codpobla char(6) default NULL,
  pobempre char(30) default NULL,
  provempre char(30) default NULL,
  nifempre char(9) default NULL,
  letraseti char(4) default NULL,
  siglasvia char(2) default NULL,
  siglaempre char(2) default NULL,
  direccion char(17) default NULL,
  numero char(4) default NULL,
  escalera char(2) default NULL,
  piso char(2) default NULL,
  puerta char(2) default NULL,
  codpos char(5) default NULL,
  poblacion char(20) default NULL,
  provincia char(15) default NULL,
  telefono char(9) default NULL,
  contacto char(100) default NULL,
  tfnocontacto char(9) default NULL,
  administracion char(5) default NULL,
  banco1 char(4) default NULL,
  oficina1 char(4) default NULL,
  dc1 char(2) default NULL,
  cuenta1 char(10) default NULL,
  banco2 char(4) default NULL,
  oficina2 char(4) default NULL,
  dc2 char(2) default NULL,
  cuenta2 char(10) default NULL,
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'hcabapu'
#

CREATE TABLE hcabapu (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien mediumint(1) unsigned NOT NULL default '0',
  obsdiari text,
  PRIMARY KEY  (numdiari,fechaent,numasien),
  KEY cl_numdiari (numdiari),
  FOREIGN KEY (numdiari) REFERENCES conta2.tiposdiario (numdiari)
) TYPE=InnoDB;



#
# Table structure for table 'hcabapu1'
#

CREATE TABLE hcabapu1 (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien mediumint(1) unsigned NOT NULL default '0',
  obsdiari text,
  PRIMARY KEY  (numdiari,fechaent,numasien),
  KEY cl_numdiari (numdiari),
  FOREIGN KEY (numdiari) REFERENCES conta2.tiposdiario (numdiari)
) TYPE=InnoDB;



#
# Table structure for table 'hlinapu'
#

CREATE TABLE hlinapu (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien mediumint(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  numdocum char(10) default NULL,
  codconce smallint(1) default NULL,
  ampconce char(30) default NULL,
  timporteD decimal(12,2) default NULL,
  codccost char(4) default NULL,
  timporteH decimal(12,2) default NULL,
  ctacontr char(10) default NULL,
  idcontab char(6) default NULL,
  punteada tinyint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (numdiari,fechaent,numasien,linliapu),
  KEY cl_numdiari (numdiari),
  KEY cl_fent (fechaent),
  KEY cl_numa (numasien),
  KEY cl_conceptos (codconce),
  KEY cl_ccostes (codccost),
  KEY cl_contrapartida (ctacontr),
  FOREIGN KEY (codconce) REFERENCES conta2.conceptos (codconce),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost),
  FOREIGN KEY (ctacontr) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (numdiari) REFERENCES conta2.tiposdiario (numdiari)
) TYPE=InnoDB;



#
# Table structure for table 'hlinapu1'
#

CREATE TABLE hlinapu1 (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien mediumint(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  numdocum char(10) default NULL,
  codconce smallint(1) default NULL,
  ampconce char(30) default NULL,
  timporteD decimal(12,2) default NULL,
  codccost char(4) default NULL,
  timporteH decimal(12,2) default NULL,
  ctacontr char(10) default NULL,
  idcontab char(6) default NULL,
  punteada tinyint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (numdiari,fechaent,numasien,linliapu),
  KEY cl_numdiari (numdiari),
  KEY cl_fent (fechaent),
  KEY cl_numa (numasien),
  KEY cl_conceptos (codconce),
  KEY cl_ccostes (codccost),
  KEY cl_contrapartida (ctacontr),
  FOREIGN KEY (codconce) REFERENCES conta2.conceptos (codconce),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost),
  FOREIGN KEY (ctacontr) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (numdiari) REFERENCES conta2.tiposdiario (numdiari)
) TYPE=InnoDB;



#
# Table structure for table 'hsaldos'
#

CREATE TABLE hsaldos (
  codmacta char(10) NOT NULL default '',
  anopsald smallint(1) NOT NULL default '0',
  mespsald tinyint(1) NOT NULL default '0',
  impmesde decimal(12,2) NOT NULL default '0.00',
  impmesha decimal(12,2) NOT NULL default '0.00',
  PRIMARY KEY  (codmacta,anopsald,mespsald),
  KEY cl_hsaldosCodmacta (codmacta),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'hsaldos1'
#

CREATE TABLE hsaldos1 (
  codmacta char(10) NOT NULL default '',
  anopsald smallint(1) NOT NULL default '0',
  mespsald tinyint(1) NOT NULL default '0',
  impmesde decimal(12,2) NOT NULL default '0.00',
  impmesha decimal(12,2) NOT NULL default '0.00',
  PRIMARY KEY  (codmacta,anopsald,mespsald),
  KEY cl_hsaldosCodmacta (codmacta),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'hsaldosanal'
#

CREATE TABLE hsaldosanal (
  codccost char(4) NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  anoccost smallint(1) NOT NULL default '0',
  mesccost tinyint(1) NOT NULL default '0',
  debccost decimal(14,2) NOT NULL default '0.00',
  habccost decimal(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (codccost,codmacta,anoccost,mesccost),
  KEY cl_hsalAnalCodmacta (codmacta),
  KEY cl_hsalAnalcodccost (codccost),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'hsaldosanal1'
#

CREATE TABLE hsaldosanal1 (
  codccost char(4) NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  anoccost smallint(1) NOT NULL default '0',
  mesccost tinyint(1) NOT NULL default '0',
  debccost decimal(14,2) NOT NULL default '0.00',
  habccost decimal(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (codccost,codmacta,anoccost,mesccost),
  KEY cl_hsalAnalCodmacta (codmacta),
  KEY cl_hsalAnalcodccost (codccost),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'linapu'
#

CREATE TABLE linapu (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '2001-01-20',
  numasien mediumint(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '0',
  numdocum varchar(10) default NULL,
  codconce smallint(1) NOT NULL default '0',
  ampconce varchar(30) default NULL,
  timporteD decimal(12,2) default NULL,
  timporteH decimal(12,2) default NULL,
  codccost varchar(4) default NULL,
  ctacontr varchar(10) default NULL,
  idcontab varchar(6) default NULL,
  punteada tinyint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (numdiari,fechaent,numasien,linliapu),
  KEY cl_cuentas2 (codmacta),
  KEY cl_contrapartida2 (ctacontr),
  KEY cl_centrocoste2 (codccost),
  KEY cl_conceptos2 (codconce),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (ctacontr) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost),
  FOREIGN KEY (codconce) REFERENCES conta2.conceptos (codconce)
) TYPE=InnoDB;



#
# Table structure for table 'linapue'
#

CREATE TABLE linapue (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '2001-01-20',
  numasien smallint(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '0',
  numdocum varchar(10) default NULL,
  codconce smallint(1) NOT NULL default '0',
  ampconce varchar(30) default NULL,
  timporteD decimal(12,2) default NULL,
  timporteH decimal(12,2) default NULL,
  codccost varchar(4) default NULL,
  ctacontr varchar(10) default NULL,
  idcontab varchar(6) default NULL,
  punteada tinyint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (numdiari,fechaent,numasien,linliapu)
) TYPE=MyISAM;



#
# Table structure for table 'linasipre'
#

CREATE TABLE linasipre (
  numaspre smallint(1) NOT NULL default '0',
  linlapre smallint(1) NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '0',
  numdocum varchar(10) default NULL,
  codconce smallint(1) NOT NULL default '0',
  ampconce varchar(30) default NULL,
  timporteD decimal(12,2) default NULL,
  timporteH decimal(12,2) default NULL,
  codccost varchar(4) default NULL,
  ctacontr varchar(10) default NULL,
  idcontab varchar(6) default NULL,
  PRIMARY KEY  (numaspre,linlapre),
  KEY cl_cuentas (codmacta),
  KEY cl_contrapartida (ctacontr),
  KEY cl_centrocoste (codccost),
  KEY cl_conceptos (codconce),
  FOREIGN KEY (codmacta) REFERENCES conta2.cuentas (codmacta) ON DELETE CASCADE,
  FOREIGN KEY (ctacontr) REFERENCES conta2.cuentas (codmacta) ON DELETE SET NULL,
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost) ON DELETE SET NULL,
  FOREIGN KEY (codconce) REFERENCES conta2.conceptos (codconce) ON DELETE CASCADE
) TYPE=InnoDB;



#
# Table structure for table 'linccost'
#

CREATE TABLE linccost (
  codccost char(4) NOT NULL default '0',
  linscost smallint(1) NOT NULL default '0',
  subccost char(4) default NULL,
  porccost decimal(5,2) NOT NULL default '0.00',
  PRIMARY KEY  (codccost,linscost),
  KEY cl_subcentro (subccost),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost) ON DELETE CASCADE,
  FOREIGN KEY (subccost) REFERENCES conta2.cabccost (codccost) ON DELETE SET NULL
) TYPE=InnoDB;



#
# Table structure for table 'linfact'
#

CREATE TABLE linfact (
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  anofaccl smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  impbascl decimal(12,2) NOT NULL default '0.00',
  codccost char(4) default NULL,
  PRIMARY KEY  (numserie,codfaccl,anofaccl,numlinea),
  KEY cl_cuentas (codtbase),
  KEY cl_ccost (codccost),
  KEY numserie (numserie,codfaccl,anofaccl),
  FOREIGN KEY (numserie, codfaccl, anofaccl) REFERENCES conta2.cabfact (numserie, codfaccl, anofaccl),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost),
  FOREIGN KEY (codtbase) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'linfact1'
#

CREATE TABLE linfact1 (
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  anofaccl smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  impbascl decimal(12,2) NOT NULL default '0.00',
  codccost char(4) default NULL,
  PRIMARY KEY  (numserie,codfaccl,anofaccl,numlinea)
) TYPE=MyISAM;



#
# Table structure for table 'linfacte'
#

CREATE TABLE linfacte (
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  anofaccl smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  impbascl decimal(12,2) NOT NULL default '0.00',
  codccost char(4) default NULL,
  PRIMARY KEY  (numserie,codfaccl,anofaccl,numlinea)
) TYPE=MyISAM;



#
# Table structure for table 'linfactprov'
#

CREATE TABLE linfactprov (
  numregis int(11) NOT NULL default '0',
  anofacpr smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  impbaspr decimal(12,2) NOT NULL default '0.00',
  codccost char(4) default NULL,
  PRIMARY KEY  (numregis,anofacpr,numlinea),
  KEY cl_cuentas (codtbase),
  KEY cl_ccost (codccost),
  KEY cl_Cabece (numregis,anofacpr),
  FOREIGN KEY (numregis, anofacpr) REFERENCES conta2.cabfactprov (numregis, anofacpr),
  FOREIGN KEY (codccost) REFERENCES conta2.cabccost (codccost),
  FOREIGN KEY (codtbase) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'linfactprov1'
#

CREATE TABLE linfactprov1 (
  numregis int(11) NOT NULL default '0',
  anofacpr smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  impbaspr decimal(12,2) NOT NULL default '0.00',
  codccost char(4) default NULL,
  PRIMARY KEY  (numregis,anofacpr,numlinea)
) TYPE=MyISAM;



#
# Table structure for table 'linfactprove'
#

CREATE TABLE linfactprove (
  numregis int(11) NOT NULL default '0',
  anofacpr smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  impbaspr decimal(12,2) NOT NULL default '0.00',
  codccost char(4) default NULL,
  PRIMARY KEY  (numregis,anofacpr,numlinea)
) TYPE=MyISAM;



#
# Table structure for table 'memoria'
#

CREATE TABLE memoria (
  codigo smallint(1) unsigned NOT NULL default '0',
  parametros tinyint(1) unsigned NOT NULL default '1',
  valor char(50) NOT NULL default '',
  descripcion char(50) default NULL,
  tipo tinyint(1) unsigned NOT NULL default '1',
  PRIMARY KEY  (codigo,parametros)
) TYPE=MyISAM;



#
# Table structure for table 'norma43'
#

CREATE TABLE norma43 (
  codigo smallint(4) NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  fecopera date NOT NULL default '0000-00-00',
  fecvalor date NOT NULL default '0000-00-00',
  importeD decimal(12,2) default NULL,
  importeH decimal(14,2) default NULL,
  concepto char(30) default NULL,
  numdocum char(10) NOT NULL default '',
  saldo decimal(14,2) default NULL,
  punteada tinyint(4) default '0',
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'parametros'
#

CREATE TABLE parametros (
  fechaini date NOT NULL default '0000-00-00',
  fechafin date NOT NULL default '0000-00-00',
  autocoste smallint(1) unsigned NOT NULL default '0',
  emitedia smallint(1) unsigned NOT NULL default '0',
  listahco smallint(1) unsigned NOT NULL default '0',
  numdiapr smallint(1) unsigned default '0',
  concefpr smallint(1) default '0',
  conceapr smallint(1) default '0',
  numdiacl smallint(1) unsigned default '0',
  concefcl smallint(1) default '0',
  conceacl smallint(1) default '0',
  limimpcl decimal(10,2) default '0.00',
  conpresu smallint(1) unsigned NOT NULL default '0',
  periodos char(2) default NULL,
  grupogto char(1) default NULL,
  grupovta char(1) default NULL,
  ctaperga varchar(10) default NULL,
  abononeg smallint(1) unsigned NOT NULL default '0',
  grupoord char(1) default NULL,
  tinumfac char(1) default NULL,
  modhcofa smallint(1) unsigned NOT NULL default '0',
  anofactu smallint(1) default NULL,
  perfactu smallint(1) default NULL,
  nctafact char(1) default NULL,
  AsienActAuto tinyint(1) NOT NULL default '0',
  codinume char(1) default NULL,
  diremail varchar(50) default NULL,
  SmtpHost varchar(50) default NULL,
  ContabilizaFact tinyint(3) unsigned zerofill NOT NULL default '000',
  SmtpUser varchar(50) default NULL,
  SmtpPass varchar(50) default NULL,
  conce43 smallint(1) default NULL,
  diario43 smallint(1) default NULL,
  constructoras tinyint(3) unsigned NOT NULL default '0',
  websoporte varchar(100) default NULL,
  mailsoporte varchar(100) default NULL,
  webversion varchar(100) default NULL,
  CCenFacturas tinyint(3) unsigned default NULL,
  Subgrupo1 varchar(10) default NULL,
  Subgrupo2 varchar(10) default NULL,
  PRIMARY KEY  (fechaini),
  KEY cl_diarios1 (numdiapr),
  KEY cl_conceptos1 (concefpr),
  KEY cl_conceptos2 (conceapr),
  KEY cl_diarios2 (numdiacl),
  KEY cl_conceptos3 (concefcl),
  KEY cl_conceptos4 (conceacl),
  KEY cl_perdygan (ctaperga),
  FOREIGN KEY (numdiapr) REFERENCES conta2.tiposdiario (numdiari),
  FOREIGN KEY (concefpr) REFERENCES conta2.conceptos (codconce),
  FOREIGN KEY (conceapr) REFERENCES conta2.conceptos (codconce),
  FOREIGN KEY (numdiacl) REFERENCES conta2.tiposdiario (numdiari),
  FOREIGN KEY (concefcl) REFERENCES conta2.conceptos (codconce),
  FOREIGN KEY (conceacl) REFERENCES conta2.conceptos (codconce),
  FOREIGN KEY (ctaperga) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'paramtesor'
#

CREATE TABLE paramtesor (
  Codigo tinyint(4) NOT NULL default '0',
  contapag tinyint(4) NOT NULL default '0',
  norma19 tinyint(4) NOT NULL default '0',
  norma32 tinyint(4) NOT NULL default '0',
  norma58 tinyint(4) NOT NULL default '0',
  cajabanco tinyint(4) default NULL,
  PRIMARY KEY  (Codigo)
) TYPE=MyISAM;



#
# Table structure for table 'presupuestos'
#

CREATE TABLE presupuestos (
  codmacta char(10) NOT NULL default '',
  anopresu smallint(6) NOT NULL default '0',
  mespresu tinyint(4) NOT NULL default '0',
  imppresu decimal(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (codmacta,anopresu,mespresu)
) TYPE=MyISAM;



#
# Table structure for table 'remesas'
#

CREATE TABLE remesas (
  codigo smallint(3) unsigned NOT NULL default '0',
  anyo smallint(3) unsigned NOT NULL default '0',
  fecremesa date default NULL,
  fecini date default NULL,
  fecfin date default NULL,
  situacion char(1) default NULL,
  codmacta varchar(10) default NULL,
  tipo tinyint(3) unsigned default NULL,
  PRIMARY KEY  (codigo,anyo)
) TYPE=MyISAM;



#
# Table structure for table 'samort'
#

CREATE TABLE samort (
  codigo tinyint(4) NOT NULL default '0',
  tipoamor tinyint(4) NOT NULL default '0',
  intcont tinyint(4) NOT NULL default '0',
  ultfecha date NOT NULL default '0000-00-00',
  condebes smallint(6) default '0',
  conhaber smallint(6) default '0',
  numdiari smallint(6) default '0',
  codiva smallint(5) unsigned default NULL,
  Preimpreso tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'sbalan'
#

CREATE TABLE sbalan (
  numbalan smallint(3) unsigned NOT NULL default '0',
  nombalan varchar(100) NOT NULL default '0',
  Descripcion text,
  Aparece tinyint(3) unsigned NOT NULL default '0',
  perdidas tinyint(3) unsigned NOT NULL default '0',
  Predeterminado tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (numbalan)
) TYPE=MyISAM;



#
# Table structure for table 'sbasin'
#

CREATE TABLE sbasin (
  codinmov smallint(6) NOT NULL default '0',
  numlinea tinyint(4) NOT NULL default '0',
  codmacta2 char(10) NOT NULL default '',
  codccost char(4) default NULL,
  porcenta decimal(5,2) NOT NULL default '0.00',
  PRIMARY KEY  (codinmov,numlinea)
) TYPE=MyISAM;



#
# Table structure for table 'scartas'
#

CREATE TABLE scartas (
  codCarta smallint(3) unsigned NOT NULL default '0',
  descarta varchar(50) default NULL,
  saludos varchar(80) default NULL,
  parrafo1 varchar(255) default NULL,
  parrafo2 varchar(255) default NULL,
  parrafo3 varchar(255) default NULL,
  desped varchar(110) default NULL,
  PRIMARY KEY  (codCarta)
) TYPE=MyISAM COMMENT='Cartas para Informes Ofertas';



#
# Table structure for table 'scobro'
#

CREATE TABLE scobro (
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  fecfaccl date NOT NULL default '0000-00-00',
  numorden smallint(1) unsigned NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '',
  codforpa smallint(6) NOT NULL default '0',
  fecvenci date NOT NULL default '0000-00-00',
  impvenci decimal(12,2) NOT NULL default '0.00',
  ctabanc1 varchar(10) NOT NULL default '',
  codbanco smallint(1) unsigned default NULL,
  codsucur smallint(1) unsigned default NULL,
  digcontr char(2) default NULL,
  cuentaba varchar(10) default NULL,
  ctabanc2 varchar(10) default NULL,
  fecultco date default '0000-00-00',
  impcobro decimal(12,2) default NULL,
  emitdocum tinyint(3) unsigned NOT NULL default '0',
  recedocu tinyint(3) unsigned NOT NULL default '0',
  contdocu tinyint(3) unsigned NOT NULL default '0',
  text33csb varchar(40) default NULL,
  text41csb varchar(40) default NULL,
  text42csb varchar(40) default NULL,
  text43csb varchar(40) default NULL,
  text51csb varchar(40) default NULL,
  text52csb varchar(40) default NULL,
  text53csb varchar(40) default NULL,
  text61csb varchar(40) default NULL,
  text62csb varchar(40) default NULL,
  text63csb varchar(40) default NULL,
  text71csb varchar(40) default NULL,
  text72csb varchar(40) default NULL,
  text73csb varchar(40) default NULL,
  text81csb varchar(40) default NULL,
  text82csb varchar(40) default NULL,
  text83csb varchar(40) default NULL,
  ultimareclamacion date default NULL,
  agente smallint(5) unsigned default NULL,
  departamento tinyint(3) unsigned default NULL,
  codrem smallint(3) unsigned default NULL,
  anyorem smallint(4) default NULL,
  siturem char(1) default NULL,
  PRIMARY KEY  (numserie,codfaccl,fecfaccl,numorden),
  KEY fp_scobro (codforpa),
  FOREIGN KEY (codforpa) REFERENCES conta2.sforpa (codforpa)
) TYPE=InnoDB;



#
# Table structure for table 'sconam'
#

CREATE TABLE sconam (
  codconam smallint(6) NOT NULL default '0',
  nomconam char(30) NOT NULL default '',
  coefimaxi decimal(5,2) NOT NULL default '0.00',
  perimaxi smallint(6) NOT NULL default '0',
  PRIMARY KEY  (codconam)
) TYPE=InnoDB;



#
# Table structure for table 'sforpa'
#

CREATE TABLE sforpa (
  codforpa smallint(6) NOT NULL default '0',
  nomforpa varchar(25) NOT NULL default '',
  tipforpa tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (codforpa)
) TYPE=InnoDB;



#
# Table structure for table 'shcocob'
#

CREATE TABLE shcocob (
  codigo int(11) NOT NULL default '0',
  numserie char(1) NOT NULL default '0',
  codfaccl int(11) NOT NULL default '0',
  fecfaccl date NOT NULL default '0000-00-00',
  numorden smallint(1) unsigned NOT NULL default '0',
  impvenci decimal(12,2) NOT NULL default '0.00',
  codmacta varchar(10) default NULL,
  nommacta varchar(35) default NULL,
  carta tinyint(4) NOT NULL default '0',
  fecreclama date NOT NULL default '0000-00-00',
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'shisin'
#

CREATE TABLE shisin (
  codinmov int(11) NOT NULL default '0',
  fechainm date NOT NULL default '0000-00-00',
  imporinm decimal(12,2) NOT NULL default '0.00',
  porcinm decimal(5,2) NOT NULL default '0.00',
  PRIMARY KEY  (codinmov,fechainm),
  KEY cl_inmov (codinmov),
  FOREIGN KEY (codinmov) REFERENCES conta2.sinmov (codinmov)
) TYPE=InnoDB;



#
# Table structure for table 'sinmov'
#

CREATE TABLE sinmov (
  codinmov int(11) NOT NULL default '0',
  codmact1 char(10) NOT NULL default '',
  nominmov char(30) NOT NULL default '',
  codprove char(10) default NULL,
  factupro char(10) default NULL,
  fechaadq date default NULL,
  codccost char(4) default NULL,
  valoradq decimal(14,2) NOT NULL default '0.00',
  codmact2 char(10) NOT NULL default '',
  codmact3 char(10) NOT NULL default '',
  conconam smallint(6) NOT NULL default '0',
  anominim smallint(6) NOT NULL default '0',
  anomaxim smallint(6) NOT NULL default '0',
  anovidas smallint(6) NOT NULL default '0',
  amortacu decimal(14,2) NOT NULL default '0.00',
  valorres decimal(14,2) NOT NULL default '0.00',
  tipoamor tinyint(4) NOT NULL default '0',
  numserie char(20) default NULL,
  fecventa date default NULL,
  impventa decimal(14,2) default NULL,
  coeficie decimal(5,2) NOT NULL default '0.00',
  situacio tinyint(4) NOT NULL default '0',
  Repartos tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (codinmov),
  KEY sinmov_Cta1 (codmact1),
  KEY sinmov_Cta2 (codmact2),
  KEY sinmov_Cta3 (codmact3),
  KEY sinmov_prove (codprove),
  FOREIGN KEY (codmact1) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (codmact2) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (codmact3) REFERENCES conta2.cuentas (codmacta),
  FOREIGN KEY (codprove) REFERENCES conta2.cuentas (codmacta)
) TYPE=InnoDB;



#
# Table structure for table 'spagop'
#

CREATE TABLE spagop (
  ctaprove varchar(10) NOT NULL default '',
  numfactu varchar(10) NOT NULL default '',
  fecfactu date NOT NULL default '0000-00-00',
  numorden smallint(1) unsigned NOT NULL default '0',
  codforpa smallint(6) NOT NULL default '0',
  fecefect date NOT NULL default '0000-00-00',
  impefect decimal(12,2) NOT NULL default '0.00',
  fecultpa date default NULL,
  imppagad decimal(12,2) default NULL,
  ctabanc1 varchar(10) NOT NULL default '',
  ctabanc2 varchar(10) default NULL,
  emitdocum tinyint(3) unsigned NOT NULL default '0',
  contdocu tinyint(3) unsigned NOT NULL default '0',
  text1csb varchar(36) default NULL,
  text2csb varchar(36) default NULL,
  PRIMARY KEY  (ctaprove,numfactu,fecfactu,numorden),
  KEY fp_spagop (codforpa),
  FOREIGN KEY (codforpa) REFERENCES conta2.sforpa (codforpa)
) TYPE=InnoDB;



#
# Table structure for table 'sparte'
#

CREATE TABLE sparte (
  normas19 tinyint(3) unsigned NOT NULL default '0',
  normas32 tinyint(3) unsigned NOT NULL default '0',
  normas58 tinyint(3) unsigned NOT NULL default '0',
  contapag tinyint(3) unsigned NOT NULL default '0'
) TYPE=MyISAM;



#
# Table structure for table 'sperdi2'
#

CREATE TABLE sperdi2 (
  NumBalan tinyint(3) unsigned NOT NULL default '0',
  Pasivo char(1) NOT NULL default '0',
  codigo tinyint(4) NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '',
  tipsaldo char(1) NOT NULL default '',
  Resta tinyint(3) unsigned NOT NULL default '0',
  PRIMARY KEY  (codigo,codmacta,Pasivo,NumBalan,Resta)
) TYPE=MyISAM;



#
# Table structure for table 'sperdid'
#

CREATE TABLE sperdid (
  NumBalan tinyint(3) unsigned NOT NULL default '0',
  Pasivo char(1) NOT NULL default '0',
  codigo tinyint(4) NOT NULL default '0',
  padre tinyint(4) default NULL,
  Orden tinyint(3) unsigned default NULL,
  tipo tinyint(4) NOT NULL default '0',
  deslinea varchar(60) NOT NULL default '',
  texlinea varchar(60) default NULL,
  formula varchar(250) default NULL,
  TienenCtas tinyint(3) unsigned NOT NULL default '0',
  Negrita tinyint(3) unsigned NOT NULL default '0',
  A_Cero tinyint(3) unsigned NOT NULL default '0',
  Pintar tinyint(3) unsigned NOT NULL default '0',
  LibroCD varchar(10) default NULL,
  PRIMARY KEY  (codigo,Pasivo,NumBalan)
) TYPE=MyISAM;



#
# Table structure for table 'stipoformapago'
#

CREATE TABLE stipoformapago (
  tipoformapago tinyint(4) NOT NULL default '0',
  descformapago varchar(25) default NULL,
  siglas varchar(5) default NULL,
  modopago smallint(5) unsigned default '0',
  modocobro smallint(5) unsigned default NULL,
  diaricli smallint(1) unsigned NOT NULL default '0',
  condecli smallint(1) NOT NULL default '0',
  conhacli smallint(1) NOT NULL default '0',
  ampdecli smallint(1) unsigned NOT NULL default '0',
  amphacli smallint(1) unsigned NOT NULL default '0',
  ctrdecli tinyint(4) NOT NULL default '0',
  ctrhacli tinyint(4) NOT NULL default '0',
  diaripro smallint(1) unsigned NOT NULL default '0',
  condepro smallint(1) NOT NULL default '0',
  conhapro smallint(1) NOT NULL default '0',
  ampdepro smallint(1) unsigned NOT NULL default '0',
  amphapro smallint(1) unsigned NOT NULL default '0',
  ctrdepro tinyint(4) NOT NULL default '0',
  ctrhapro tinyint(4) NOT NULL default '0',
  PRIMARY KEY  (tipoformapago)
) TYPE=MyISAM;



#
# Table structure for table 'tipoamortizacion'
#

CREATE TABLE tipoamortizacion (
  tipoamor tinyint(4) NOT NULL default '0',
  desctipoamor char(10) NOT NULL default '',
  PRIMARY KEY  (tipoamor)
) TYPE=MyISAM;



#
# Table structure for table 'tipoconceptos'
#

CREATE TABLE tipoconceptos (
  tipoconce tinyint(1) NOT NULL default '0',
  desctipo char(20) NOT NULL default '0',
  PRIMARY KEY  (tipoconce)
) TYPE=MyISAM;



#
# Table structure for table 'tipomemoria'
#

CREATE TABLE tipomemoria (
  codigo tinyint(1) unsigned NOT NULL default '1',
  descripcion char(50) default NULL,
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'tiposdiario'
#

CREATE TABLE tiposdiario (
  numdiari smallint(1) unsigned NOT NULL default '0',
  desdiari char(30) NOT NULL default '',
  PRIMARY KEY  (numdiari)
) TYPE=InnoDB;



#
# Table structure for table 'tiposituacion'
#

CREATE TABLE tiposituacion (
  situacio tinyint(4) NOT NULL default '0',
  descsituacion char(10) NOT NULL default '',
  PRIMARY KEY  (situacio)
) TYPE=MyISAM;



#
# Table structure for table 'tiposiva'
#

CREATE TABLE tiposiva (
  codigiva tinyint(1) unsigned NOT NULL default '0',
  nombriva char(15) NOT NULL default '',
  tipodiva tinyint(1) unsigned NOT NULL default '0',
  porceiva decimal(4,2) NOT NULL default '0.00',
  porcerec decimal(4,2) default NULL,
  cuentare char(10) NOT NULL default '',
  cuentarr char(10) NOT NULL default '',
  cuentaso char(10) NOT NULL default '',
  cuentasr char(10) NOT NULL default '',
  cuentasn char(10) NOT NULL default '',
  PRIMARY KEY  (codigiva),
  KEY cl_repercutido (cuentare),
  KEY cl_repercutidorec (cuentarr),
  KEY cl_soportado (cuentaso),
  KEY cl_soportadore (cuentasr),
  KEY cl_soportadoNded (cuentasn),
  FOREIGN KEY (cuentare) REFERENCES conta2.cuentas (codmacta) ON DELETE CASCADE,
  FOREIGN KEY (cuentarr) REFERENCES conta2.cuentas (codmacta) ON DELETE CASCADE,
  FOREIGN KEY (cuentaso) REFERENCES conta2.cuentas (codmacta) ON DELETE CASCADE,
  FOREIGN KEY (cuentasr) REFERENCES conta2.cuentas (codmacta) ON DELETE CASCADE,
  FOREIGN KEY (cuentasn) REFERENCES conta2.cuentas (codmacta) ON DELETE CASCADE
) TYPE=InnoDB;



#
# Table structure for table 'tmp347'
#

CREATE TABLE tmp347 (
  codusu smallint(1) unsigned NOT NULL default '0',
  cliprov tinyint(4) NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  nif varchar(15) default '',
  importe decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,cliprov,cta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpactualizar'
#

CREATE TABLE tmpactualizar (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien smallint(1) unsigned NOT NULL default '0',
  codusu smallint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (numdiari,fechaent,numasien,codusu)
) TYPE=MyISAM;



#
# Table structure for table 'tmpactualizarerror'
#

CREATE TABLE tmpactualizarerror (
  numdiari smallint(1) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien smallint(1) unsigned NOT NULL default '0',
  codusu smallint(1) unsigned NOT NULL default '0',
  Error varchar(200) default NULL,
  PRIMARY KEY  (numdiari,fechaent,numasien,codusu)
) TYPE=MyISAM;



#
# Table structure for table 'tmpbussinmov'
#

CREATE TABLE tmpbussinmov (
  codmacta varchar(10) NOT NULL default '',
  titulo varchar(30) default NULL,
  PRIMARY KEY  (codmacta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpcierre'
#

CREATE TABLE tmpcierre (
  Importe decimal(12,2) default NULL,
  cta char(10) NOT NULL default '',
  PRIMARY KEY  (cta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpcierre1'
#

CREATE TABLE tmpcierre1 (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  nomcta varchar(30) NOT NULL default '0',
  acumPerD decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,cta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpconext'
#

CREATE TABLE tmpconext (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  numdiari smallint(1) unsigned NOT NULL default '0',
  Pos int(10) unsigned NOT NULL default '0',
  fechaent date NOT NULL default '2001-01-20',
  numasien smallint(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  nomdocum varchar(10) default NULL,
  ampconce varchar(30) default NULL,
  timporteD decimal(10,2) default NULL,
  timporteH decimal(10,2) default NULL,
  saldo decimal(10,2) default NULL,
  Punteada char(2) default NULL,
  contra varchar(10) default NULL,
  ccost varchar(4) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'tmpconextcab'
#

CREATE TABLE tmpconextcab (
  codusu smallint(1) unsigned NOT NULL default '0',
  cuenta varchar(80) NOT NULL default '',
  fechini varchar(10) default NULL,
  fechfin varchar(10) default NULL,
  acumantD decimal(14,2) default NULL,
  acumantH decimal(14,2) default NULL,
  acumantT decimal(14,2) default NULL,
  acumperD decimal(14,2) default NULL,
  acumperH decimal(14,2) default NULL,
  acumperT decimal(14,2) default NULL,
  acumtotD decimal(14,2) default NULL,
  acumtotH decimal(14,2) default NULL,
  acumtotT decimal(14,2) default NULL,
  cta varchar(10) NOT NULL default '',
  PRIMARY KEY  (codusu,cta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpctaexpcc'
#

CREATE TABLE tmpctaexpcc (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  codccost varchar(4) NOT NULL default '',
  PRIMARY KEY  (codccost,codusu,cta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpctaexplotacioncierre'
#

CREATE TABLE tmpctaexplotacioncierre (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  acumPerD decimal(14,2) default NULL,
  acumPerH decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,cta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpdiarresum'
#

CREATE TABLE tmpdiarresum (
  codusu smallint(5) unsigned NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  Debe decimal(14,2) default NULL,
  Haber decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,codmacta)
) TYPE=MyISAM;



#
# Table structure for table 'tmpfaclin'
#

CREATE TABLE tmpfaclin (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo smallint(1) unsigned NOT NULL default '0',
  Numfac varchar(12) default NULL,
  Fecha varchar(10) default NULL,
  cta varchar(10) default NULL,
  Cliente varchar(30) default NULL,
  NIF varchar(12) default NULL,
  Imponible decimal(14,2) default NULL,
  IVA varchar(5) default NULL,
  ImpIVA decimal(14,2) default NULL,
  Total decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'tmpimpbalance'
#

CREATE TABLE tmpimpbalance (
  codusu smallint(4) NOT NULL default '0',
  Pasivo char(1) NOT NULL default '',
  codigo smallint(6) NOT NULL default '0',
  descripcion char(30) default NULL,
  linea char(60) default NULL,
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  negrita tinyint(4) default NULL,
  orden smallint(6) NOT NULL default '0',
  PRIMARY KEY  (codigo,codusu,Pasivo)
) TYPE=MyISAM;



#
# Table structure for table 'tmpliqiva'
#

CREATE TABLE tmpliqiva (
  codusu smallint(1) unsigned NOT NULL default '0',
  iva decimal(14,2) NOT NULL default '0.00',
  PRIMARY KEY  (codusu,iva)
) TYPE=MyISAM;



#
# Table structure for table 'tmppresu1'
#

CREATE TABLE tmppresu1 (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo int(11) NOT NULL default '0',
  cta varchar(10) default NULL,
  titulo varchar(30) default NULL,
  ano smallint(6) NOT NULL default '0',
  mes tinyint(4) NOT NULL default '0',
  Importe decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'tmpresumenivafac'
#

CREATE TABLE tmpresumenivafac (
  codusu smallint(1) unsigned NOT NULL default '0',
  orden smallint(1) unsigned NOT NULL default '0',
  IVA varchar(10) default NULL,
  TotalIVA decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,orden)
) TYPE=MyISAM;



#
# Table structure for table 'tmpsimula'
#

CREATE TABLE tmpsimula (
  codusu smallint(6) NOT NULL default '0',
  codinmov int(6) NOT NULL default '0',
  codconam smallint(6) NOT NULL default '0',
  totalamor decimal(12,2) NOT NULL default '0.00',
  PRIMARY KEY  (codusu,codinmov)
) TYPE=MyISAM;



#
# Table structure for table 'tmpsperdi'
#

CREATE TABLE tmpsperdi (
  codusu tinyint(3) unsigned NOT NULL default '0',
  pasivo char(1) NOT NULL default '0',
  codigo tinyint(4) NOT NULL default '0',
  importe decimal(12,2) default NULL,
  PRIMARY KEY  (codusu,pasivo,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'zbloqueos'
#

CREATE TABLE zbloqueos (
  codusu smallint(1) unsigned NOT NULL default '0',
  tabla char(20) NOT NULL default '',
  clave char(30) NOT NULL default '',
  PRIMARY KEY  (tabla,clave)
) TYPE=MyISAM;

