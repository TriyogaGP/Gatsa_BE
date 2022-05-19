const multer = require('multer');
const path = require('path');

const storage = multer.diskStorage({
    destination: (req, file, callBack) => {
        const { body } = req;
        const { jenis, nama, id } = body
        const location = (jenis === 'images') ? './src/public/images/' : (jenis === 'excel') ? './src/public/excel/' : './src/public/pdf/'
        callBack(null, location)     // './public/images/' directory name where save the file
    },
    filename: (req, file, callBack) => {
        const { body } = req;
        const { jenis, nama, id } = body
        let ubahNama = nama.replace(' ', '_');
        callBack(null, ubahNama + path.extname(file.originalname))
    }
})

const uploadFile = multer({
    storage: storage
}).any();

module.exports = {
    uploadFile
}