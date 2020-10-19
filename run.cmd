@echo off

java -Xdebug -Xrunjdwp:transport=dt_socket,address=8453,server=y,suspend=n -classpath spire.doc.free-2.7.3.jar;target\mailmerge-1.0-SNAPSHOT.jar ca.ulaval.pul.SimpleMailMerge

