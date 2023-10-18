const xlsx = require('node-xlsx');
const path = require('path');
const fs = require('fs');
const _ = require('lodash');

const filePath = path.join(__dirname, 'sources', '20231018.xls');

function readXLSFile() {
    const workSheetsFromFile = xlsx.parse(filePath);

    workSheetsFromFile.forEach((workSheet, sheetId) => {
        const temp = {};
        const existedRole = [];
        if (workSheet.name === 'porta原版 ') {
            workSheet.data.forEach((row, rowNumber) => {
                if (rowNumber === 0) return;

                const companyName = _.get(row, 4);
                const roleName = _.get(row, 1);
                const id = _.get(row, 2);
                const name = _.get(row, 3);

                if (_.isEmpty(roleName)) return;

                if (!_.has(temp, companyName)) {
                    _.set(temp, companyName, {});
                }

                if (!_.includes(existedRole, roleName)) {
                    existedRole.push(roleName);
                }

                const roles = _.get(temp, companyName);
                if (!_.has(roles, roleName)) {
                    _.set(roles, roleName, []);
                }

                roles[roleName].push({ id, name });
            });

            const newData = [['公司']];
            _.forEach(existedRole, role => {
                const item = newData[0];
                item.push(role + "(工号)");
                item.push(role + "(姓名)");
            })

            _.forEach(temp, (obj, companyName) => {
                let maxLength = 0;
                _.forEach(existedRole, role => {
                    const currentLength = _.size(_.get(obj, role));
                    if (currentLength > maxLength) {
                        maxLength = currentLength;
                    }
                })

                for (let i = 0; i < maxLength; ++i) {
                    const item = [companyName];
                    _.forEach(existedRole, role => {
                        const roleObj = _.get(obj, role, []);
                        const o = _.get(roleObj, i, {});
                        item.push(o.id || '');
                        item.push(o.name || '');
                    })
                    newData.push(item);
                }
            });

            const newFilePath = path.join(__dirname, 'transformed-excel-file.xlsx');
            const buffer = xlsx.build([{ name: 'Transformed Data', data: newData }]);

            fs.writeFileSync(newFilePath, buffer);
        }
    });
}

readXLSFile();
