const multer = require('multer')
const path = require('path')
const fs = require('fs');

const storage = multer.diskStorage({
    destination: (req, file, callBack) => {
        const { body } = req;
        const { nama, id, nomor_induk } = body
        const path_dir = path.join(__dirname, '../public/pdf/' + id + '-' + nomor_induk);
        // console.log(path_dir);
        // process.exit()
        if (!fs.existsSync(path_dir)) {
            fs.mkdirSync(path_dir, 0777);
        };
        callBack(null, path_dir)     // './public/images/' directory name where save the file
    },
    filename: (req, file, callBack) => {
        const { body } = req;
        const { nama, id, nomor_induk } = body
        callBack(null, nama + path.extname(file.originalname))
    }
})

const uploadBerkas = multer({
    storage: storage
}).any();

module.exports = {
    uploadBerkas
}