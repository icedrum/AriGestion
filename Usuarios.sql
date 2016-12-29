# MySQL-Front Dump 2.5
#
# Host: localhost   Database: Usuarios
# --------------------------------------------------------
# Server version 3.23.53-max-nt


#
# Table structure for table 'appmenus'
#

CREATE TABLE appmenus (
  aplicacion varchar(15) default '0',
  Name varchar(100) default '0',
  caption varchar(100) default '0',
  indice tinyint(3) default '0',
  padre smallint(3) unsigned default '0',
  orden tinyint(3) unsigned default NULL,
  Contador smallint(5) unsigned default NULL
) TYPE=MyISAM;



#
# Table structure for table 'appmenususuario'
#

CREATE TABLE appmenususuario (
  aplicacion varchar(15) NOT NULL default '0',
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo smallint(3) unsigned NOT NULL default '0',
  tag varchar(100) default '0',
  PRIMARY KEY  (aplicacion,codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'empresas'
#

CREATE TABLE empresas (
  codempre tinyint(4) NOT NULL default '0',
  nomempre char(50) NOT NULL default '',
  nomresum char(15) NOT NULL default '',
  Usuario char(20) default NULL,
  Pass char(20) default NULL,
  Conta char(10) default NULL,
  PRIMARY KEY  (codempre)
) TYPE=MyISAM COMMENT='Empresas en el sistema';



#
# Table structure for table 'empresasumi'
#

CREATE TABLE empresasumi (
  codempre tinyint(4) NOT NULL default '0',
  nomempre char(50) NOT NULL default '',
  nomresum char(15) NOT NULL default '',
  Usuario char(20) default NULL,
  Pass char(20) default NULL,
  Sumi char(10) default NULL,
  PRIMARY KEY  (codempre)
) TYPE=MyISAM COMMENT='Empresas en el sistema';



#
# Table structure for table 'pcs'
#

CREATE TABLE pcs (
  codpc smallint(5) unsigned NOT NULL default '0',
  nompc char(30) default NULL,
  PRIMARY KEY  (codpc)
) TYPE=MyISAM;



#
# Table structure for table 'usuarioempresa'
#

CREATE TABLE usuarioempresa (
  codusu smallint(1) unsigned NOT NULL default '0',
  codempre smallint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (codusu,codempre)
) TYPE=MyISAM;



#
# Table structure for table 'usuarioempresasumi'
#

CREATE TABLE usuarioempresasumi (
  codusu smallint(1) unsigned default '0',
  codempre smallint(1) unsigned default '0'
) TYPE=MyISAM;



#
# Table structure for table 'usuarios'
#

CREATE TABLE usuarios (
  codusu smallint(1) unsigned NOT NULL default '0',
  nomusu char(30) NOT NULL default '',
  dirfich char(50) default NULL,
  nivelusu tinyint(1) NOT NULL default '-1',
  login char(20) NOT NULL default '',
  passwordpropio char(20) NOT NULL default '',
  nivelusuges tinyint(4) NOT NULL default '-1',
  nivelariges tinyint(4) NOT NULL default '-1',
  nivelsumi tinyint(4) default '-1',
  PRIMARY KEY  (codusu)
) TYPE=MyISAM;



#
# Table structure for table 'vbloqbd'
#

CREATE TABLE vbloqbd (
  codusu smallint(5) unsigned NOT NULL default '0',
  conta char(30) default NULL,
  PRIMARY KEY  (codusu)
) TYPE=InnoDB;



#
# Table structure for table 'wasientos'
#

CREATE TABLE wasientos (
  Lugar tinyint(3) unsigned NOT NULL default '0',
  codigo smallint(1) NOT NULL default '0',
  numdiari char(10) default NULL,
  fechaent char(10) default NULL,
  numasien char(11) default NULL,
  linliapu char(10) default NULL,
  codmacta char(10) default NULL,
  iddebhab char(1) default NULL,
  numdocum char(10) default NULL,
  codconce char(10) default NULL,
  ampconce char(30) default NULL,
  ctacontr char(10) default NULL,
  codccost char(10) default NULL,
  idcontab char(6) default NULL,
  importe char(12) default NULL,
  nommacta1 char(30) default NULL,
  nommacta2 char(30) default NULL,
  Observa char(250) default NULL,
  correcto tinyint(3) unsigned default NULL,
  PRIMARY KEY  (codigo,Lugar)
) TYPE=MyISAM;



#
# Table structure for table 'wcabfact'
#

CREATE TABLE wcabfact (
  lugar tinyint(4) NOT NULL default '0',
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  fecfaccl date NOT NULL default '0000-00-00',
  codmacta varchar(10) NOT NULL default '',
  anofaccl smallint(6) NOT NULL default '0',
  confaccl varchar(15) default NULL,
  pi1faccl decimal(6,2) default NULL,
  pi2faccl decimal(6,2) default NULL,
  pi3faccl decimal(6,2) default NULL,
  pr1faccl decimal(6,2) default NULL,
  pr2faccl decimal(6,2) default NULL,
  pr3faccl decimal(6,2) default NULL,
  tp1faccl tinyint(1) unsigned NOT NULL default '0',
  tp2faccl tinyint(3) unsigned default NULL,
  tp3faccl tinyint(3) unsigned default NULL,
  intracom tinyint(3) unsigned NOT NULL default '0',
  ba1faccl decimal(12,2) NOT NULL default '0.00',
  ba2faccl decimal(12,2) default NULL,
  ba3faccl decimal(12,2) default NULL,
  ti1faccl decimal(12,2) default NULL,
  ti2faccl decimal(12,2) default NULL,
  ti3faccl decimal(12,2) default NULL,
  tr1faccl decimal(12,2) default NULL,
  tr2faccl decimal(12,2) default NULL,
  tr3faccl decimal(12,2) default NULL,
  totfaccl decimal(14,2) default NULL,
  retfaccl decimal(6,2) default NULL,
  trefaccl decimal(12,2) default NULL,
  cuereten varchar(10) default NULL,
  nommacta varchar(30) default NULL,
  fecliqcl date default NULL,
  SeContabiliza char(1) default NULL,
  correcto tinyint(3) unsigned default NULL,
  PRIMARY KEY  (lugar,numserie,codfaccl,anofaccl)
) TYPE=MyISAM;



#
# Table structure for table 'wcabfactp'
#

CREATE TABLE wcabfactp (
  lugar tinyint(4) NOT NULL default '0',
  codfaccl int(11) NOT NULL default '0',
  numfaccl varchar(10) NOT NULL default '',
  fecfaccl date NOT NULL default '0000-00-00',
  codmacta varchar(10) NOT NULL default '',
  anofaccl smallint(6) NOT NULL default '0',
  confaccl varchar(15) default NULL,
  pi1faccl decimal(6,2) default NULL,
  pi2faccl decimal(6,2) default NULL,
  pi3faccl decimal(6,2) default NULL,
  pr1faccl decimal(6,2) default NULL,
  pr2faccl decimal(6,2) default NULL,
  pr3faccl decimal(6,2) default NULL,
  retfaccl decimal(6,2) default NULL,
  frefacpr date default NULL,
  tp1faccl tinyint(1) unsigned NOT NULL default '0',
  tp2faccl tinyint(3) unsigned default NULL,
  tp3faccl tinyint(3) unsigned default NULL,
  idffaccl tinyint(4) NOT NULL default '0',
  ba1faccl decimal(12,2) NOT NULL default '0.00',
  ba2faccl decimal(12,2) default NULL,
  ba3faccl decimal(12,2) default NULL,
  ti1faccl decimal(12,2) default NULL,
  ti2faccl decimal(12,2) default NULL,
  ti3faccl decimal(12,2) default NULL,
  tr1faccl decimal(12,2) default NULL,
  tr2faccl decimal(12,2) default NULL,
  tr3faccl decimal(12,2) default NULL,
  totfaccl decimal(14,2) default NULL,
  trefaccl decimal(12,2) default NULL,
  cuereten varchar(10) default NULL,
  nommacta varchar(30) default NULL,
  fecliqcl date default NULL,
  SeContabiliza char(1) default NULL,
  correcto tinyint(3) unsigned default NULL,
  PRIMARY KEY  (lugar,codfaccl,anofaccl)
) TYPE=MyISAM;



#
# Table structure for table 'wlinfact'
#

CREATE TABLE wlinfact (
  lugar tinyint(4) NOT NULL default '0',
  numserie char(1) NOT NULL default '',
  codfaccl int(11) NOT NULL default '0',
  anofaccl smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  impbascl decimal(12,2) NOT NULL default '0.00',
  codccost char(4) default NULL,
  correcto tinyint(3) unsigned default NULL,
  PRIMARY KEY  (lugar,numserie,codfaccl,anofaccl,numlinea)
) TYPE=MyISAM;



#
# Table structure for table 'wlinfactp'
#

CREATE TABLE wlinfactp (
  lugar tinyint(4) NOT NULL default '0',
  codfaccl int(11) NOT NULL default '0',
  anofaccl smallint(6) NOT NULL default '0',
  numlinea smallint(6) NOT NULL default '0',
  codtbase char(10) NOT NULL default '',
  codccost char(4) default NULL,
  impbascl decimal(12,2) NOT NULL default '0.00',
  correcto tinyint(3) unsigned default NULL,
  PRIMARY KEY  (lugar,codfaccl,anofaccl,numlinea)
) TYPE=MyISAM;



#
# Table structure for table 'wnorma43'
#

CREATE TABLE wnorma43 (
  codusu smallint(4) NOT NULL default '0',
  Orden smallint(5) unsigned NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  fecopera date NOT NULL default '0000-00-00',
  fecvalor date NOT NULL default '0000-00-00',
  importeD decimal(12,2) default NULL,
  importeH decimal(14,2) default NULL,
  concepto char(30) default NULL,
  numdocum char(10) NOT NULL default '',
  saldo decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,Orden,codmacta)
) TYPE=MyISAM;



#
# Table structure for table 'ypruebalogo'
#

CREATE TABLE ypruebalogo (
  id tinyint(3) unsigned NOT NULL default '0',
  Imagen blob,
  PRIMARY KEY  (id)
) TYPE=MyISAM;



#
# Table structure for table 'z347'
#

CREATE TABLE z347 (
  codusu smallint(1) unsigned NOT NULL default '0',
  cliprov tinyint(4) NOT NULL default '0',
  nif varchar(15) NOT NULL default '',
  importe decimal(14,2) default NULL,
  razosoci varchar(30) default NULL,
  dirdatos varchar(30) default '',
  codposta varchar(6) default '',
  despobla varchar(30) default '',
  PRIMARY KEY  (codusu,cliprov,nif)
) TYPE=MyISAM;



#
# Table structure for table 'z347carta'
#

CREATE TABLE z347carta (
  codusu smallint(1) unsigned NOT NULL default '0',
  nif varchar(15) NOT NULL default '',
  razosoci varchar(30) default NULL,
  dirdatos varchar(30) default '',
  codposta varchar(6) default '',
  despobla varchar(30) default '',
  otralineadir varchar(40) default NULL,
  saludos varchar(100) default NULL,
  parrafo1 varchar(255) default NULL,
  parrafo2 varchar(255) default NULL,
  parrafo3 varchar(255) default NULL,
  parrafo4 varchar(255) default NULL,
  parrafo5 varchar(255) default NULL,
  despedida varchar(100) default NULL,
  contacto varchar(50) default NULL,
  Asunto varchar(100) default NULL,
  Referencia varchar(30) default NULL,
  PRIMARY KEY  (codusu)
) TYPE=MyISAM;



#
# Table structure for table 'zasipre'
#

CREATE TABLE zasipre (
  codusu smallint(1) unsigned NOT NULL default '0',
  numaspre smallint(1) NOT NULL default '0',
  nomaspre varchar(40) NOT NULL default '',
  linlapre smallint(1) NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '0',
  nommacta varchar(30) NOT NULL default '',
  ampconce varchar(30) default NULL,
  timporteD decimal(12,2) default NULL,
  timporteH decimal(12,2) default NULL,
  codccost varchar(4) default NULL,
  PRIMARY KEY  (codusu,numaspre,linlapre)
) TYPE=MyISAM;



#
# Table structure for table 'zcabccexplo'
#

CREATE TABLE zcabccexplo (
  codusu smallint(1) unsigned NOT NULL default '0',
  codccost char(4) NOT NULL default '',
  codmacta char(10) NOT NULL default '',
  nommacta char(30) NOT NULL default '',
  nomccost char(30) NOT NULL default '',
  acumD decimal(12,2) default NULL,
  TieneAcum char(1) default NULL,
  acumH decimal(12,2) default NULL,
  acumS decimal(12,2) default NULL,
  totD decimal(12,2) default NULL,
  totH decimal(12,2) default NULL,
  totS decimal(12,2) default NULL,
  PRIMARY KEY  (codusu,codmacta,codccost)
) TYPE=MyISAM;



#
# Table structure for table 'zcertifiva'
#

CREATE TABLE zcertifiva (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo int(1) NOT NULL default '0',
  factura varchar(11) default NULL,
  fecha varchar(10) default NULL,
  destino varchar(30) default NULL,
  Importe decimal(12,2) default NULL,
  tipoiva tinyint(4) default NULL,
  iva varchar(5) default NULL,
  pais varchar(15) default NULL,
  nif varchar(15) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'zconceptos'
#

CREATE TABLE zconceptos (
  codusu smallint(1) unsigned NOT NULL default '0',
  codconce char(5) NOT NULL default '0',
  nomconce char(30) NOT NULL default '0',
  tipoconce char(15) NOT NULL default '0',
  PRIMARY KEY  (codusu,codconce)
) TYPE=MyISAM;



#
# Table structure for table 'zctaexpcc'
#

CREATE TABLE zctaexpcc (
  codusu smallint(1) unsigned NOT NULL default '0',
  codccost char(4) NOT NULL default '0',
  nomccost char(30) NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  nommacta char(30) default NULL,
  acumD decimal(12,2) default NULL,
  acumH decimal(12,2) default NULL,
  perid decimal(12,2) default NULL,
  periH decimal(12,2) default NULL,
  postD decimal(12,2) default NULL,
  postH decimal(12,2) default NULL,
  saldoD decimal(12,2) default NULL,
  saldoH decimal(12,2) default NULL,
  PRIMARY KEY  (codusu,codccost,codmacta)
) TYPE=MyISAM;



#
# Table structure for table 'zcuentas'
#

CREATE TABLE zcuentas (
  codusu smallint(1) unsigned NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '0',
  nommacta varchar(30) NOT NULL default '0',
  razosoci varchar(30) default '',
  dirdatos varchar(30) default '',
  codposta varchar(6) default '',
  despobla varchar(30) default '',
  nifdatos varchar(15) default '',
  apudirec char(1) default NULL,
  model347 tinyint(3) default '0',
  PRIMARY KEY  (codusu,codmacta)
) TYPE=MyISAM;



#
# Table structure for table 'zdiapendact'
#

CREATE TABLE zdiapendact (
  codusu smallint(6) NOT NULL default '0',
  numdiari smallint(1) unsigned NOT NULL default '0',
  desdiari varchar(30) NOT NULL default '',
  fechaent date NOT NULL default '0000-00-00',
  numasien mediumint(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '0',
  nommacta varchar(30) NOT NULL default '0',
  numdocum varchar(10) default NULL,
  ampconce varchar(30) default NULL,
  timporteD decimal(12,2) default NULL,
  timporteH decimal(12,2) default NULL,
  codccost varchar(4) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'zdirioresum'
#

CREATE TABLE zdirioresum (
  codusu smallint(5) unsigned NOT NULL default '0',
  clave smallint(5) unsigned NOT NULL default '0',
  fecha char(10) default NULL,
  asiento smallint(6) default NULL,
  cuenta char(30) default NULL,
  titulo char(30) default NULL,
  concepto char(30) default NULL,
  debe decimal(14,2) default NULL,
  haber decimal(14,2) default NULL,
  PRIMARY KEY  (clave,codusu)
) TYPE=MyISAM;



#
# Table structure for table 'zentrefechas'
#

CREATE TABLE zentrefechas (
  codusu smallint(5) unsigned NOT NULL default '0',
  codigo smallint(6) NOT NULL default '0',
  codccost char(4) default NULL,
  nomccost char(30) default NULL,
  conconam smallint(6) default NULL,
  nomconam char(30) default NULL,
  codinmov smallint(6) NOT NULL default '0',
  nominmov char(30) NOT NULL default '',
  fechaadq char(10) default NULL,
  valoradq decimal(12,2) default NULL,
  amortacu decimal(14,2) default NULL,
  fecventa date default NULL,
  impventa decimal(14,2) default NULL,
  impperiodo decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'zestadinmo1'
#

CREATE TABLE zestadinmo1 (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo smallint(1) NOT NULL default '0',
  codconam smallint(6) NOT NULL default '0',
  nomconam char(30) NOT NULL default '',
  codinmov int(6) NOT NULL default '0',
  nominmov char(30) NOT NULL default '',
  tipoamor char(1) default NULL,
  porcenta char(6) default NULL,
  codprove char(10) default NULL,
  fechaadq char(10) default NULL,
  valoradq decimal(12,2) default NULL,
  amortacu decimal(12,5) default NULL,
  fecventa char(10) default NULL,
  impventa decimal(12,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'zexplocomp'
#

CREATE TABLE zexplocomp (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  cuenta varchar(30) NOT NULL default '',
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  activo tinyint(4) default NULL,
  PRIMARY KEY  (codusu,cta)
) TYPE=MyISAM;



#
# Table structure for table 'zexplocompimpre'
#

CREATE TABLE zexplocompimpre (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo smallint(6) NOT NULL default '0',
  cta varchar(10) default NULL,
  cuenta varchar(30) default NULL,
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  2cta varchar(10) default NULL,
  2cuenta varchar(30) default NULL,
  2importe1 decimal(14,2) default NULL,
  2importe2 decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'zfichainmo'
#

CREATE TABLE zfichainmo (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo smallint(1) NOT NULL default '0',
  codinmov int(6) NOT NULL default '0',
  nominmov char(30) NOT NULL default '',
  fechaadq char(10) default NULL,
  valoradq decimal(12,2) default NULL,
  fechaamor char(10) default NULL,
  Importe decimal(12,2) default NULL,
  porcenta decimal(6,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'zhistoapu'
#

CREATE TABLE zhistoapu (
  codusu smallint(1) unsigned NOT NULL default '0',
  numdiari smallint(1) unsigned NOT NULL default '0',
  desdiari char(30) NOT NULL default '',
  fechaent date NOT NULL default '0000-00-00',
  numasien mediumint(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  codmacta char(10) NOT NULL default '',
  nommacta char(30) default NULL,
  numdocum char(10) default NULL,
  ampconce char(30) default NULL,
  timporteD decimal(12,2) default NULL,
  timporteH decimal(12,2) default NULL,
  codccost char(4) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'zificheros'
#

CREATE TABLE zificheros (
  codigo smallint(1) unsigned NOT NULL default '0',
  path varchar(50) NOT NULL default '0',
  nombre varchar(11) NOT NULL default '0',
  PRIMARY KEY  (codigo)
) TYPE=MyISAM;



#
# Table structure for table 'zlinccexplo'
#

CREATE TABLE zlinccexplo (
  codusu smallint(1) unsigned NOT NULL default '0',
  codccost varchar(4) NOT NULL default '',
  codmacta varchar(10) NOT NULL default '',
  linapu smallint(6) NOT NULL default '0',
  docum varchar(10) NOT NULL default '',
  fechaent date NOT NULL default '0000-00-00',
  ampconce varchar(30) default NULL,
  perD decimal(12,2) default NULL,
  perH decimal(12,2) default NULL,
  saldo decimal(12,2) default NULL,
  ctactra varchar(10) default NULL,
  desctra varchar(20) default NULL,
  PRIMARY KEY  (codusu,codmacta,codccost,linapu,fechaent)
) TYPE=MyISAM;



#
# Table structure for table 'zliquidaiva'
#

CREATE TABLE zliquidaiva (
  codusu smallint(1) unsigned NOT NULL default '0',
  iva decimal(14,2) NOT NULL default '0.00',
  bases decimal(14,2) default NULL,
  ivas decimal(14,2) default NULL,
  codempre tinyint(3) unsigned NOT NULL default '0',
  periodo tinyint(3) unsigned NOT NULL default '0',
  ano smallint(3) unsigned NOT NULL default '0',
  cliente tinyint(3) unsigned NOT NULL default '1',
  PRIMARY KEY  (codusu,iva,ano,codempre,periodo,cliente)
) TYPE=MyISAM;



#
# Table structure for table 'zmemoria'
#

CREATE TABLE zmemoria (
  codusu smallint(5) unsigned NOT NULL default '0',
  codigo smallint(6) NOT NULL default '0',
  parame tinyint(4) NOT NULL default '0',
  descripcion char(50) default NULL,
  valortexto char(50) default NULL,
  texto2 char(16) default NULL,
  valornumero decimal(12,2) default NULL,
  PRIMARY KEY  (codusu,codigo,parame)
) TYPE=MyISAM;



#
# Table structure for table 'zpendientes'
#

CREATE TABLE zpendientes (
  codusu smallint(1) unsigned NOT NULL default '0',
  serie_cta varchar(10) NOT NULL default '',
  factura varchar(10) NOT NULL default '',
  fecha date NOT NULL default '0000-00-00',
  numorden smallint(1) unsigned NOT NULL default '0',
  codforpa smallint(6) NOT NULL default '0',
  nomforpa varchar(25) NOT NULL default '',
  codmacta varchar(10) default NULL,
  nombre varchar(30) NOT NULL default '',
  fecVto date NOT NULL default '0000-00-00',
  importe decimal(12,2) NOT NULL default '0.00',
  pag_cob decimal(12,2) NOT NULL default '0.00',
  vencido tinyint(4) NOT NULL default '0',
  PRIMARY KEY  (codusu,serie_cta,factura,fecha,numorden)
) TYPE=MyISAM;



#
# Table structure for table 'zsaldoscc'
#

CREATE TABLE zsaldoscc (
  codusu smallint(1) unsigned NOT NULL default '0',
  codccost char(4) NOT NULL default '0',
  nomccost char(30) NOT NULL default '0',
  ano smallint(1) NOT NULL default '0',
  mes tinyint(1) NOT NULL default '0',
  impmesde decimal(12,2) NOT NULL default '0.00',
  impmesha decimal(12,2) NOT NULL default '0.00',
  PRIMARY KEY  (ano,codusu,codccost,mes)
) TYPE=MyISAM;



#
# Table structure for table 'zsimulainm'
#

CREATE TABLE zsimulainm (
  codusu smallint(5) unsigned NOT NULL default '0',
  codigo smallint(6) NOT NULL default '0',
  conconam smallint(6) default NULL,
  nomconam char(30) default NULL,
  codinmov int(6) NOT NULL default '0',
  nominmov char(30) NOT NULL default '',
  fechaadq char(10) default NULL,
  valoradq decimal(14,2) default NULL,
  amortacu decimal(14,2) default NULL,
  totalamor decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'ztesoreriacomun'
#

CREATE TABLE ztesoreriacomun (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo smallint(1) unsigned NOT NULL default '0',
  texto1 varchar(35) default NULL,
  texto2 varchar(35) default NULL,
  texto3 varchar(35) default NULL,
  texto4 varchar(35) default NULL,
  texto5 varchar(35) default NULL,
  texto6 varchar(35) default NULL,
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  fecha1 date default NULL,
  fecha2 date default NULL,
  fecha3 date default NULL,
  observa1 varchar(255) default NULL,
  observa2 varchar(255) default NULL,
  opcion tinyint(4) default '0',
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'ztiposdiario'
#

CREATE TABLE ztiposdiario (
  codusu smallint(1) unsigned NOT NULL default '0',
  numdiari smallint(1) unsigned NOT NULL default '0',
  desdiari char(30) NOT NULL default '',
  PRIMARY KEY  (codusu,numdiari)
) TYPE=MyISAM;



#
# Table structure for table 'ztmpbalanceconsolidado'
#

CREATE TABLE ztmpbalanceconsolidado (
  codempre smallint(6) NOT NULL default '0',
  nomempre varchar(30) NOT NULL default '',
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  nomcta varchar(30) NOT NULL default '0',
  aperturaD decimal(14,2) default NULL,
  aperturaH decimal(14,2) default NULL,
  acumAntD decimal(14,2) default NULL,
  acumAntH decimal(14,2) default NULL,
  acumPerD decimal(14,2) default NULL,
  acumPerH decimal(14,2) default NULL,
  TotalD decimal(14,2) default NULL,
  TotalH decimal(14,2) default NULL,
  PRIMARY KEY  (codempre,codusu,cta)
) TYPE=MyISAM;



#
# Table structure for table 'ztmpbalancesumas'
#

CREATE TABLE ztmpbalancesumas (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  nomcta varchar(30) NOT NULL default '0',
  aperturaD decimal(14,2) default NULL,
  aperturaH decimal(14,2) default NULL,
  acumAntD decimal(14,2) default NULL,
  acumAntH decimal(14,2) default NULL,
  acumPerD decimal(14,2) default NULL,
  acumPerH decimal(14,2) default NULL,
  TotalD decimal(14,2) default NULL,
  TotalH decimal(14,2) default NULL
) TYPE=MyISAM;



#
# Table structure for table 'ztmpconext'
#

CREATE TABLE ztmpconext (
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
# Table structure for table 'ztmpconextcab'
#

CREATE TABLE ztmpconextcab (
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
# Table structure for table 'ztmpctaexplotacion'
#

CREATE TABLE ztmpctaexplotacion (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) default NULL,
  Contador int(10) unsigned NOT NULL default '0',
  nomcta varchar(30) NOT NULL default '0',
  acumAntD decimal(14,2) default NULL,
  acumAntH decimal(14,2) default NULL,
  acumPerD decimal(14,2) default NULL,
  acumPerH decimal(14,2) default NULL,
  TotalD decimal(14,2) default NULL,
  TotalH decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,Contador)
) TYPE=MyISAM;



#
# Table structure for table 'ztmpctaexplotacionc'
#

CREATE TABLE ztmpctaexplotacionc (
  codusu smallint(1) unsigned NOT NULL default '0',
  cta varchar(10) NOT NULL default '',
  codempre smallint(1) unsigned NOT NULL default '0',
  empresa varchar(50) default NULL,
  nomcta varchar(30) NOT NULL default '0',
  acumAntD decimal(14,2) default NULL,
  acumAntH decimal(14,2) default NULL,
  acumPerD decimal(14,2) default NULL,
  acumPerH decimal(14,2) default NULL,
  TotalD decimal(14,2) default NULL,
  TotalH decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,cta,codempre)
) TYPE=MyISAM;



#
# Table structure for table 'ztmpfaclin'
#

CREATE TABLE ztmpfaclin (
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
  retencion decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'ztmpfaclinprov'
#

CREATE TABLE ztmpfaclinprov (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo smallint(1) unsigned NOT NULL default '0',
  Numfac varchar(12) default NULL,
  FechaFac varchar(10) default NULL,
  FechaCon varchar(10) default NULL,
  cta varchar(10) default NULL,
  Cliente varchar(30) default NULL,
  NIF varchar(12) default NULL,
  Imponible decimal(14,2) default NULL,
  IVA varchar(5) default NULL,
  ImpIVA decimal(14,2) default NULL,
  Total decimal(14,2) default NULL,
  retencion decimal(14,2) default NULL,
  NoDeducible char(2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'ztmpimpbalan'
#

CREATE TABLE ztmpimpbalan (
  codusu smallint(1) unsigned NOT NULL default '0',
  Pasivo char(1) NOT NULL default '',
  codigo smallint(6) NOT NULL default '0',
  descripcion varchar(60) default NULL,
  linea varchar(100) default NULL,
  importe1 decimal(14,2) default NULL,
  importe2 decimal(14,2) default NULL,
  negrita tinyint(4) default NULL,
  LibroCD varchar(6) default NULL,
  PRIMARY KEY  (codigo,codusu,Pasivo)
) TYPE=MyISAM;



#
# Table structure for table 'ztmplibrodiario'
#

CREATE TABLE ztmplibrodiario (
  codusu smallint(6) NOT NULL default '0',
  fechaent date NOT NULL default '0000-00-00',
  numasien int(1) unsigned NOT NULL default '0',
  linliapu smallint(1) unsigned NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '',
  nommacta varchar(30) NOT NULL default '0',
  numdocum varchar(10) default NULL,
  ampconce varchar(30) default NULL,
  debe decimal(14,2) default '0.00',
  haber decimal(14,2) default '0.00',
  PRIMARY KEY  (codusu,numasien,linliapu)
) TYPE=MyISAM;



#
# Table structure for table 'ztmppresu1'
#

CREATE TABLE ztmppresu1 (
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
# Table structure for table 'ztmppresu2'
#

CREATE TABLE ztmppresu2 (
  codusu smallint(1) unsigned NOT NULL default '0',
  codigo int(11) NOT NULL default '0',
  cta varchar(10) default NULL,
  titulo varchar(30) default NULL,
  mes tinyint(4) default '0',
  Presupuesto decimal(14,2) default NULL,
  realizado decimal(6,2) default NULL,
  PRIMARY KEY  (codusu,codigo)
) TYPE=MyISAM;



#
# Table structure for table 'ztmpresumenivafac'
#

CREATE TABLE ztmpresumenivafac (
  codusu smallint(1) unsigned NOT NULL default '0',
  orden smallint(1) unsigned NOT NULL default '0',
  IVA varchar(10) default NULL,
  TotalIVA decimal(14,2) default NULL,
  sumabases decimal(14,2) default NULL,
  PRIMARY KEY  (codusu,orden)
) TYPE=MyISAM;



#
# Table structure for table 'ztotalctaconce'
#

CREATE TABLE ztotalctaconce (
  codusu smallint(6) NOT NULL default '0',
  codmacta varchar(10) NOT NULL default '',
  nommacta varchar(30) default NULL,
  nifdatos varchar(15) default NULL,
  fechaent date NOT NULL default '0000-00-00',
  timporteD decimal(14,2) default NULL,
  timporteH decimal(12,2) default NULL,
  codconce smallint(1) default NULL
) TYPE=MyISAM;

