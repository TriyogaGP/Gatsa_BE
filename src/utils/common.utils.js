const multipleColumnSetWhere = (object) => {
  if (typeof object !== 'object') {
      throw new Error('Invalid input');
  }

  const keys = Object.keys(object);
  const values = Object.values(object);

  columnSet = keys.map(key => `${key} = ?`).join(' AND ');

  return {
      columnSet,
      values
  }
}

const multipleColumnSet = (object) => {
  if (typeof object !== 'object') {
      throw new Error('Invalid input');
  }

  const keys = Object.keys(object);
  const values = Object.values(object);

  columnSet = keys.map(key => `${key} = ?`).join(', ');

  return {
      columnSet,
      values
  }
}

const multipleWhere = (object) => {
  if (typeof object !== 'object') {
      throw new Error('Invalid input');
  }
  let OR = object.or;
  let AND = object.and;  
  let dataKumpulOR, dataKumpulAND, dataKumpul;

  if(Object.entries(OR).length){
    for (const [key, value] of Object.entries(OR)) {
      if(key == 'and') delete OR[key];
    }
    
    const keys = Object.keys(OR);
    const values = Object.values(OR);
    columnSet = (values.length > 1) ? '('+keys.map(key => `${key} = ?`).join(' OR ')+')' : keys.map(key => `${key} = ?`).join(' OR ')

    if(!Object.entries(AND).length){
      return {
        columnSet,
        values
      }
    }else{
      dataKumpulOR = {
        columnSet,
        values
      }
    }
  }
  
  if(Object.entries(AND).length){
    for (const [key, value] of Object.entries(AND)) {
      if(key == 'or') delete AND[key];
    }
    
    const keys = Object.keys(AND);
    const values = Object.values(AND);
  
    columnSet = (values.length > 1) ? '('+keys.map(key => `${key} = ?`).join(' AND ')+')' : keys.map(key => `${key} = ?`).join(' AND ')
    
    if(!Object.entries(OR).length){
      return {
        columnSet,
        values
      }
    }else{
      dataKumpulAND = {
        columnSet,
        values
      }
    }
  }

  if(Object.entries(OR).length && Object.entries(AND).length){
    let operator = 
      (dataKumpulOR.values.length > dataKumpulAND.values.length) ? 'OR' : 
      (dataKumpulOR.values.length < dataKumpulAND.values.length) ? 'AND' : 'SAME'
    let columnSet = 
      (operator == 'OR') ? dataKumpulOR.columnSet+' AND '+dataKumpulAND.columnSet :
      (operator == 'AND') ? dataKumpulAND.columnSet+' OR '+dataKumpulOR.columnSet : dataKumpulOR.columnSet+' AND '+dataKumpulAND.columnSet
    let values = 
      (operator == 'OR') ? dataKumpulOR.values.concat(dataKumpulAND.values) :
      (operator == 'AND') ? dataKumpulAND.values.concat(dataKumpulOR.values) : dataKumpulOR.values.concat(dataKumpulAND.values)

    return {
      columnSet,
      values
    }
  }

}

const multipleValueSet = (object) => {
  if (typeof object !== 'object') {
      throw new Error('Invalid input');
  }

  const values = Object.values(object);
  
  return { values }
}

module.exports = {
  multipleColumnSet,
  multipleColumnSetWhere,
  multipleWhere,
  multipleValueSet
}