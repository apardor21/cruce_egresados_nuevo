<?php
declare(strict_types=1);
function h($v):string{return htmlspecialchars((string)$v,ENT_QUOTES|ENT_SUBSTITUTE,'UTF-8');}
function qi(string $s):string{return '`'.str_replace('`','``',$s).'`';}
function normdoc($v):string{$s=trim((string)$v);if(preg_match('/^\d+\.0+$/',$s))$s=explode('.',$s)[0];return mb_strtoupper(preg_replace('/[\s.\-]+/u','',$s)??$s,'UTF-8');}
