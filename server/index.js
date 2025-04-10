const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const mongoose = require('mongoose');

const { getRectangleFromExcel, getRange } = require('./utils/parser');

const { AudsModel, KafsModel } = require('./models/index');

const app = express();
const port = 5000;

app.use(express.json());
app.use(cors());
app.use(express.static(path.resolve(__dirname, './dist/')));

const start = async () => {
    try { 
        await mongoose.connect('mongodb://database:27017/schedule-viewer');
        // await mongoose.connect('mongodb://localhost:27017/schedule-viewer');

        app.listen(port, () => {
            console.log(`Server started on http://localhost:${port}`);
        });
    } catch (e) {
        console.log('Ошибка подключения к БД:', e);
    }
};

// удаление кафдеры с ее аудиториями или удаление аудитории
app.delete('/api/delete', async (req, res) => {
    try {
        const { audId, kafId } = req.body;

        if (audId) {
            await AudsModel.deleteOne({ _id: audId });
            return res.status(200).json({ message: 'Аудитория удалена' });
        }

        if (kafId) {
            const wantedKaf = await KafsModel.findOne({ _id: kafId });
            const wantedAuds = wantedKaf.audsIds;

            for (let i = 0; i <= wantedAuds.length; i += 1) {
                await AudsModel.deleteOne({ _id: wantedAuds[i] });
            }

            await KafsModel.deleteOne({ _id: kafId });

            return res.status(200).json({ message: 'Кафедра удалена' });
        }
    } catch (error) {
        console.log(error);
        return res.status(500).json({ message: 'Произошла непредвиденная ошибка' });
    }
});

// создание кафедры
app.post('/api/create_kaf', async (req, res) => {
    try {
        const { title } = req.body;

        const createdKaf = KafsModel.create({
            title,
        });

        return res.status(201).json(createdKaf);
    } catch (error) {
        console.log(error);
        return res.status(500).json({ message: 'Произошла непредвиденная ошибка' });
    }
});

// добавление аудиторий к кафедре по ее id
app.post('/api/add_auds_to_kaf', async (req, res) => {
    try {
        const { audsTitles, parentKafId } = req.body;

        const createdAudsIds = [];
        for (let i = 0; i < audsTitles.length; i += 1) {
            const createdAud = await AudsModel.create({ title: audsTitles[i] });
            createdAudsIds.push(createdAud._id);
        }

        await KafsModel.updateOne({ _id: parentKafId }, { audsIds: createdAudsIds });

        return res.status(201).json({ message: 'Аудитории созданы и добавлены' });
    } catch (error) {
        console.log(error);
        return res.status(500).json({ message: 'Произошла непредвиденная ошибка' });
    }
});

app.get('/api/fetch_auds', async (req, res) => {
    try {
        const auds = await AudsModel.find({}).lean();
        if (!auds?.length) return res.status(404).json({ message: 'Аудиторий не найдено' });
        return res.status(200).json(auds);
    } catch (e) {
        console.log(e);
        return res.status(500).json({ message: 'Непредвиденная ошибка на сервере' });
    }
});

// найти аудитории по id кафедры
app.get('/api/find_by_kaf', async (req, res) => {
    try {
        const { kafId } = req.query; // { kafId: kj2niu1nrijwenjkgnsdkjng }

        const wantedKaf = await KafsModel.findOne({ _id: kafId }).populate({ path: 'audsIds' });
        if (!wantedKaf)
            return res.status(404).json({ message: `Кафедра с ID = ${kafId} не найдена` });

        const wantedAuds = wantedKaf.audsIds;
        if (!wantedAuds.length)
            return res
                .status(404)
                .json({ message: `За кафедрой с ID = ${kafId} аудитории не закреплены` });

        return res.status(200).json(wantedAuds);
    } catch (e) {
        console.log(e);
        return res.status(500).json({ message: 'Непредвиденная ошибка на сервере' });
    }
});

// получить список кафедр
app.get('/api/get_kafs', async (req, res) => {
    try {
        const { populate } = req.query;

        let kafs = [];

        if (populate) kafs = await KafsModel.find({}).populate({ path: 'audsIds' });
        else kafs = await KafsModel.find({});

        if (!kafs.length) return res.status(404).json({ message: 'Кафедры не найдены' });

        return res.status(200).json(kafs);
    } catch (e) {
        console.log(e);
        return res.status(500).json({ message: 'Непредвиденная ошибка на сервере' });
    }
});

app.get('/api/groups', (req, res) => {
    const { dir } = req.query;

    try {
        // Проверяем, является ли путь директорией
        const stats = fs.statSync(dir);
        
        if (!stats.isDirectory()) {
            return res.status(400).json({ message: 'Указанный путь не является директорией' });
        }
        
        const files = fs.readdirSync(dir);
        // Фильтруем только Excel-файлы и исключаем временные файлы
        const excelFiles = files.filter(file => 
            (file.endsWith('.xlsx') || file.endsWith('.xls')) && !file.includes('~')
        );
        
        const title = excelFiles.map((file) => path.basename(file).split('.')[0]);

        return res.status(200).json(title);
    } catch (e) {
        console.log('ошибка', e);
        return res.status(500).json({ message: 'Директория не найдена или произошла ошибка при чтении' });
    }
});

app.get('/api/subjects', (req, res) => {
    const { group, workDir } = req.query;
    
    try {
        // Используем workDir, если он предоставлен, иначе используем локальный путь
        const filePath = workDir 
            ? `${workDir}${workDir.endsWith('/') || workDir.endsWith('\\') ? '' : '/'}${group}.xlsx` 
            : path.resolve(__dirname, `./files/${group}.xlsx`);
            
        const subjects = getRange(filePath, 'A39:O60');
        const filteredSubjects = subjects.filter(s => s.abbr?.length > 1 && s.abbr?.length <= 4);
        
        return res.status(200).json(filteredSubjects);
    } catch (error) {
        console.log('Ошибка при чтении файла:', error);
        return res.status(500).json({ message: 'Файл не найден или произошла ошибка при чтении' });
    }
});

app.get('/api/schedule', async (req, res) => {
    try {
        const { workDir, group, kafId } = req.query;

        // Исправляем формирование пути к файлу, добавляя проверку на завершающий слеш
        const filePath = workDir 
            ? `${workDir}${workDir.endsWith('/') || workDir.endsWith('\\') ? '' : '/'}${group}.xlsx` 
            : path.resolve(__dirname, `./files/${group}.xlsx`);
            
        const schedule = getRectangleFromExcel(filePath, 'D6:Y34');

        if (kafId) {
            const thisKaf = await KafsModel.findOne({ _id: kafId }).populate({ path: 'audsIds' });
            const audsTitle = thisKaf.audsIds.map((aud) => aud.title);

            const filteredByKaf = [];
            for (let i = 0; i < schedule.length; i += 1) {
                for (let j = 0; j < audsTitle.length; j += 1) {
                    const hello = schedule[i].jobs.map(
                        (job) =>
                            job.includes(audsTitle[j]) &&
                            !job.includes('самоподготовка') &&
                            !job.includes('хозяйственный день'),
                    );

                    if (
                        hello.some((str) => str) &&
                        !filteredByKaf.find((day) => day.date === schedule[i].date)
                    ) {
                        filteredByKaf.push(schedule[i]);
                    }
                }
            }

            return res.status(200).json(filteredByKaf);
        }

        return res.status(200).json(schedule);
    } catch (error) {
        console.log(error);
        return res.status(500).json({ message: 'Произошла непредвиденная ошибка' });
    }
});

app.get('/api/today', async (req, res) => {
    try { 
        const { workDir } = req.query;
        console.log('API /api/today called with workDir:', workDir);

        const groupsSchedule = [];
        let cnt = 0;

        console.log('Reading directory:', path.resolve(__dirname, workDir));
        const schedule = fs
            .readdirSync(path.resolve(__dirname, workDir))
            .filter((file) => !file.includes('~'));
        
        console.log('Found files:', schedule);

        schedule.forEach((file) => {
            console.log('Processing file:', file);
            const filePath = path.join(path.resolve(__dirname, workDir), file);
            console.log('Full file path:', filePath);
            
            const data = getRectangleFromExcel(filePath, 'D6:Y34');
            console.log(`Extracted ${data.length} days from ${file}`);
            groupsSchedule.push(data);
        });

        console.log('Total groups processed:', groupsSchedule.length);

        const result = [];
        const filteredGroups = groupsSchedule
            .map((group) => { 
                const filteredDays = group.filter((day) => {
                    // Устанавливаем конкретную дату 11.04.2025 (месяцы 0-индексированные)
                    const today = new Date(2025, 3, 11).setHours(0, 0, 0, 0);
                    const date = new Date(day.date).setHours(0, 0, 0, 0);
                    return today === date;
                });
                console.log(`Found ${filteredDays.length} matching days for group`);
                return filteredDays;
            });
            
        filteredGroups.forEach((group) => {
            if (group.length > 0) {
                console.log('Adding group to results:', schedule[cnt]?.split('.')[0]);
                group[0].groupName = schedule[cnt]?.split('.')[0];
                result.push(group[0]);
            }
            cnt += 1;
        });

        console.log('Final result count:', result.length);
        console.log('Result data:', JSON.stringify(result, null, 2));

        return res.status(200).json(result);
    } catch (e) {
        console.log('Error in /api/today endpoint:', e);        
        return res.status(500).json({message: 'Произошла ошибка'})
    }
});

app.get('/*', (req, res) => res.sendFile(path.resolve(__dirname, './dist/index.html')));

start();
