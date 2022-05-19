module.exports = {
    response: (res, data, statusCode) => {
        res.header("Access-Control-Allow-Origin", "*");
        res.status(statusCode)
        res.json(data)
        res.end()
    },
    OK: (res, data = null, message = null) => {
      res.status(200);
      res.json({
        status: 'SUCCESS',
        message: message || 'SUCCESS',
        data,
      });
    },
    NO_CONTENT: res => {
      res.status(204);
      res.json({
        status: 'SUCCESS',
        message: 'SUCCESS',
        data: null,
      });
  
      return res;
    },
    CREATED: (res, data = null, message = null) => {
      res.status(201);
      res.json({
        status: 'SUCCESS',
        message: message || 'SUCCESS',
        data,
      });
  
      return res;
    },
    NOT_FOUND: (res, message) => {
      res.status(404);
      res.json({
        status: 'ERROR',
        message: message || 'Requested data not found',
        data: null,
      });
  
      return res;
    },
    UNAUTHORIZED: (res, message) => {
      res.status(401);
      res.json({
        status: 'ERROR',
        message: message || 'Unauthorized',
        data: null,
      });
  
      return res;
    },
    ERROR: (res, message, data = null) => {
      res.status(500);
      res.json({
        status: 'ERROR',
        message: message || 'An error occurred trying to process your request',
        data,
      });
  
      return res;
    },
    UNPROCESSABLE: (res, message, data = null) => {
      res.status(422);
      res.json({
        status: 'ERROR',
        message: message || 'Unprocessable entity',
        data,
      });
  
      return res;
    },
  };
  