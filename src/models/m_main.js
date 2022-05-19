const { result } = require('lodash');
const _ = require('lodash');
const query = require('../config/database');
const { multipleColumnSetWhere, multipleColumnSet, multipleWhere, multipleValueSet } = require('../utils/common.utils');

class LoginModel {

  findWhere = async (prototype = {}) => {
    let sql;
    let attr_includes = [], attr = [], wadah_attr = [], order_attr_includes = [], order_attr = [], tampung;
    let multipleData;
    if(Object.entries(prototype).length){
              
      // untuk attributes table utama
      if(prototype.table){
        if(!prototype.attributes.length){
          order_attr.push(`${prototype.table}.*`);
        }else{
          prototype.attributes.map((value, key) => {
            tampung = _.split(value, ' ');
            tampung = (tampung.length > 1) ? _.last(tampung) : tampung[0];
            order_attr.push(tampung);
          });
        }
      }

      //  untuk atributes di include
      if(prototype.include.joinTable){
        if(!prototype.include.attributes.length) { 
          order_attr_includes.push(`${prototype.include.joinTable}.*`);
        }else{
          prototype.include.attributes.map((value, key) => {
            tampung = _.split(value, ' ');
            tampung = (tampung.length > 1) ? _.last(tampung) : tampung[0];
            order_attr_includes.push(tampung);
          });
        }
      }

      if(!Object.entries(prototype.include).length){
        if(!prototype.attributes.length){
          attr.push(`${prototype.table}.*`);
          wadah_attr = attr.join(', ');
        }else{
          prototype.attributes.map((value, key) => {
            tampung = `${prototype.table}.${value}`;
            attr.push(tampung);
          });
          wadah_attr = attr.join(', ');
        }
      }else{
        if(!prototype.include.attributes.length) { 
          if(!prototype.attributes.length){
            attr.push(`${prototype.table}.*`);
          }else{
            prototype.attributes.map((value, key) => {
              tampung = `${prototype.table}.${value}`;
              attr.push(tampung);
            });
          }
          attr_includes.push(`${prototype.include.joinTable}.*`);
          wadah_attr = _.concat(attr, attr_includes).join(', ');
        }else{
          if(!prototype.attributes.length){
            attr.push(`${prototype.table}.*`);
          }else{
            prototype.attributes.map((value, key) => {
              tampung = `${prototype.table}.${value}`;
              attr.push(tampung);
            });
          }
          prototype.include.attributes.map((value, key) => {
            tampung = `${prototype.include.joinTable}.${value}`;
            attr_includes.push(tampung);
          });
          wadah_attr = _.concat(attr,attr_includes).join(', ');
        }
      }

      sql = `SELECT ROW_NUMBER() OVER(ORDER BY ${prototype.on} ASC) AS item_no, ${wadah_attr} FROM ${prototype.table}`;

      if(prototype.on && prototype.include.joinOn) { sql += ` INNER JOIN ${prototype.include.joinTable} ON ${prototype.on} = ${prototype.include.joinOn}`; }
      if(Object.entries(prototype.where.or).length || Object.entries(prototype.where.and).length) { 
        // console.log(multipleData);
        multipleData = multipleWhere(prototype.where);
        sql += ` WHERE ${multipleData.columnSet}`;
      }else{
        multipleData = {
          columnSet: '',
          values: []
        }
      }
      if(prototype.orderBy) { sql += ` ORDER BY ${!prototype.orderByValue ? 'item_no' : `${prototype.orderByValue}`} ${prototype.orderBy}`; }
    }

    const result = await query(sql, [...multipleData.values]);
    // return (result.length > 1 ) ? result : (typeof result[0] == 'undefined') ? {} : result[0];
    return result
  }

  findWilayah = async (params = {}) => {
    const { values } = multipleValueSet(params.search);
   
    let sql;
    if(params.field == 'provinsi' || params.field == 'kabkotaOnly'){
      sql = `SELECT kode AS value, nama AS label, kode_pos FROM wilayah WHERE CHAR_LENGTH(kode)= ? ORDER BY nama`;
    }else{
      sql = `SELECT kode AS value, nama AS label, kode_pos FROM wilayah WHERE LEFT(kode,?)= ? AND CHAR_LENGTH(kode)= ? ORDER BY nama`;
    }
    const result = await query(sql, [...values]);
    return result
  }

  findKelas = async (params = {}, where = {}) => {
    const { columnSet, values } = multipleWhere(where);
    // console.log(columnSet, values)
    // process.exit()
    let sql;
    if(params.kelas == 'ALL'){
      sql = `SELECT CONCAT_WS('-', kelas, number) as value, CONCAT_WS('-', kelas, number) as label FROM kelas WHERE ${columnSet} ORDER BY id_kelas ASC`;
    }else{
      sql = `SELECT CONCAT_WS('-', kelas, number) as value, CONCAT_WS('-', kelas, number) as label FROM kelas WHERE ${columnSet} ORDER BY id_kelas ASC`;
    }
    const result = await query(sql, [...values]);
    return result
  }

  update = async (params = {}, table, where = {}) => {
    let sql;
    const { columnSet, values } = multipleColumnSet(params);
    const { columnSet: item, values: value } = multipleColumnSetWhere(where);

    sql = `UPDATE ${table} SET ${columnSet} WHERE ${item}`;
    
    const result = await query(sql, values.concat(value));
    
    return result;
  }

  insert = async (params = {}, table) => {
    let sql;
    const { columnSet, values } = multipleColumnSet(params);

    sql = `INSERT INTO ${table} SET ${columnSet}`;
    
    const result = await query(sql, [...values]);
    const affectedRows = result ? true : false;
    return affectedRows;
  }

  delete = async (params = {}, table) => {
    let sql;
    const { columnSet, values } = multipleColumnSet(params);

    sql = `DELETE FROM ${table} WHERE ${columnSet}`;
    
    const result = await query(sql, [...values]);
    const affectedRows = result ? true : false;
    return affectedRows;
  }

}

module.exports = new LoginModel;