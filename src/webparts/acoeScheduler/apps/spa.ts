import * as $ from 'jquery';
import '@progress/kendo-ui';
import { sp } from '@pnp/sp/presets/all';

export class ModelState extends kendo.data.ObservableObject {
    protected collapsed: string[] = [];

    constructor() {
        super();
    }
}

export class SPA {
    protected static schedulerOptions: kendo.ui.SchedulerOptions;
    protected static scheduler: kendo.ui.Scheduler;
    private static instance: SPA;

    constructor() {}

    public static getInstance(calendars: string[]): SPA {
        this.schedulerOptions = null;
        this.schedulerOptions = null;
        const state = new ModelState();

        var dataArray = [];
        var currentDate = new Date();
        currentDate.setDate(currentDate.getDate() - 14);    //back limited to 2 weeks

        const getLocalDate = (isoString: string): Date => {
            let [year, month, date] = (isoString).split('T')[0].split('-').map(Number);
            let [hour, minute, second] = (isoString).split('T')[1].split(':').map(Number);

            return new Date (year, month, date, hour, minute);
        };

        // convert Outlook recurrence to iCal rules
        const recurrenceParser = (recurrenceData: string): string => {
            let rruleStr = '', frequency = '', byMonth = '', byDay = 'BYDAY=', bySetPos = '', byMonthDay = '', interval = '', until = '', count = '';

            if (recurrenceData.indexOf('rule') != -1) 
            {
                var xmlStr = $.parseXML(recurrenceData);
                let ruleType = $(xmlStr).find('repeat').children()[0].localName;
                let posValue;

                switch (ruleType) {
                    case 'yearly':
                    case 'yearlyByDay':
                        frequency = 'FREQ=yearly;';
    
                        // Repeat Rule #1
                        if ($(xmlStr).find('yearly').length > 0) {
                            byMonthDay = 'BYMONTHDAY=' + $(xmlStr).find('yearly').attr('day') + ';';
                            byMonth = 'BYMONTH=' + $(xmlStr).find('yearly').attr('month') + ';';
    
                            rruleStr = frequency + byMonth + byMonthDay;
                        }
    
                        // Repeat Rule #2
                        if ($(xmlStr).find('yearlyByDay').length > 0) {
                            switch ($(xmlStr).find('yearlyByDay').attr('weekdayOfMonth')) {
                                case 'first': posValue = 1; break;
                                case 'second': posValue = 2; break;
                                case 'third': posValue = 3; break;
                                case 'fourth': posValue = 4; break;
                                default: posValue = -1;
                            }
                            if ($(xmlStr).find('yearlyByDay').attr('su') !== undefined) byDay += 'SU;';
                            if ($(xmlStr).find('yearlyByDay').attr('mo') !== undefined) byDay += 'MO;';
                            if ($(xmlStr).find('yearlyByDay').attr('tu') !== undefined) byDay += 'TU;';
                            if ($(xmlStr).find('yearlyByDay').attr('we') !== undefined) byDay += 'WE;';
                            if ($(xmlStr).find('yearlyByDay').attr('th') !== undefined) byDay += 'TH;';
                            if ($(xmlStr).find('yearlyByDay').attr('fr') !== undefined) byDay += 'FR;';
                            if ($(xmlStr).find('yearlyByDay').attr('sa') !== undefined) byDay += 'SA;';
                            if ($(xmlStr).find('yearlyByDay').attr('day') !== undefined) byDay += 'SU,MO,TU,WE,TH,FR,SA;';
                            if ($(xmlStr).find('yearlyByDay').attr('weekday') !== undefined) byDay += 'MO,TU,WE,TH,FR;';
                            bySetPos = 'BYSETPOS=' + posValue + ';';
                            byMonth = 'BYMONTH=' + $(xmlStr).find('yearlyByDay').attr('month') + ';';
    
                            rruleStr = frequency + byDay + bySetPos + byMonth;
                        }
                        break;
    
                    case 'monthly':
                    case 'monthlyByDay':
                        frequency = 'FREQ=monthly;';
    
                        // Repeat Rule #1
                        if ($(xmlStr).find('monthly').length > 0) {
                            interval = 'INTERVAL=' + $(xmlStr).find('monthly').attr('monthFrequency') + ';';
                            byMonthDay = 'BYMONTHDAY=' + $(xmlStr).find('monthly').attr('day') + ';';
    
                            rruleStr = frequency + byMonthDay + interval;
                        }
    
                        // Repeat Rule #2
                        if ($(xmlStr).find('monthlyByDay').length > 0) {
                            interval = 'INTERVAL=' + $(xmlStr).find('monthlyByDay').attr('monthFrequency') + ';';
                            switch ($(xmlStr).find('monthlyByDay').attr('weekdayOfMonth')) {
                                case 'first': posValue = 1; break;
                                case 'second': posValue = 2; break;
                                case 'third': posValue = 3; break;
                                case 'fourth': posValue = 4; break;
                                default: posValue = -1;
                            }
                            if ($(xmlStr).find('monthlyByDay').attr('su') !== undefined) byDay += 'SU;';
                            if ($(xmlStr).find('monthlyByDay').attr('mo') !== undefined) byDay += 'MO;';
                            if ($(xmlStr).find('monthlyByDay').attr('tu') !== undefined) byDay += 'TU;';
                            if ($(xmlStr).find('monthlyByDay').attr('we') !== undefined) byDay += 'WE;';
                            if ($(xmlStr).find('monthlyByDay').attr('th') !== undefined) byDay += 'TH;';
                            if ($(xmlStr).find('monthlyByDay').attr('fr') !== undefined) byDay += 'FR;';
                            if ($(xmlStr).find('monthlyByDay').attr('sa') !== undefined) byDay += 'SA;';
                            if ($(xmlStr).find('monthlyByDay').attr('day') !== undefined) byDay += 'SU,MO,TU,WE,TH,FR,SA;';
                            if ($(xmlStr).find('monthlyByDay').attr('weekday') !== undefined) byDay += 'MO,TU,WE,TH,FR;';
                            bySetPos = 'BYSETPOS=' + posValue + ';';
    
                            rruleStr = frequency + bySetPos + byDay + interval;
                        }
                        break;
    
                    case 'weekly':
                        frequency = 'FREQ=weekly;';
                        if ($(xmlStr).find('weekly').attr('su') !== undefined) byDay += 'SU';
                        if ($(xmlStr).find('weekly').attr('mo') !== undefined) byDay += (byDay === 'BYDAY=') ? 'MO' : ',MO';
                        if ($(xmlStr).find('weekly').attr('tu') !== undefined) byDay += (byDay === 'BYDAY=') ? 'TU' : ',TU';
                        if ($(xmlStr).find('weekly').attr('we') !== undefined) byDay += (byDay === 'BYDAY=') ? 'WE' : ',WE';
                        if ($(xmlStr).find('weekly').attr('th') !== undefined) byDay += (byDay === 'BYDAY=') ? 'TH' : ',TH';
                        if ($(xmlStr).find('weekly').attr('fr') !== undefined) byDay += (byDay === 'BYDAY=') ? 'FR' : ',FR';
                        if ($(xmlStr).find('weekly').attr('sa') !== undefined) byDay += (byDay === 'BYDAY=') ? 'SA' : ',SA';
                        byDay += ';';
                        interval = 'INTERVAL=' + $(xmlStr).find('weekly').attr('weekFrequency') + ';';
    
                        rruleStr = frequency + byDay + interval;
                        break;
                    case 'daily':
                        frequency = 'FREQ=daily;';
                        if ($(xmlStr).find('daily').attr('dayFrequency') != undefined) interval = 'INTERVAL=' + $(xmlStr).find('daily').attr('dayFrequency') + ';';
    
                        rruleStr = frequency + interval;
                        break;
                }    
            }

            // End portion of the rule
            if ($(xmlStr).find('repeatInstances').length > 0) count = 'COUNT=' + $(xmlStr).find('repeatInstances').text() + ';';
            if ($(xmlStr).find('windowEnd').length > 0) count = 'UNTIL=' + $(xmlStr).find('windowEnd').text() + ';';
            rruleStr += count + until + 'WKST=SU;';

            return rruleStr.toUpperCase();
        };

        const getCalendarData = async (calArray) => {
            const recurse = (next: any) => {
                next.getNext().then(nestedResponse => {
                    dataArray = [...dataArray, ...nestedResponse.results];
                    if (nestedResponse.hasNext) recurse(nestedResponse);
                    else console.log(dataArray);
                });
            };

            // Assign a team ID based on the order of entry in the property panel
            let teamId = 0;
            for (const site of calArray) {
                sp.setup({
                    sp: {
                        baseUrl: site.URL,
                        headers: {'Accept': 'application/json;odata=nometadata'}
                    }
                });

                await sp.web.lists.getByTitle(site.Title).items.select('Id','Title','Location','EventDate','EndDate','Description','fAllDayEvent','fRecurrence','Category','RecurrenceData').top(1000).filter("EndDate ge datetime'" + currentDate.toISOString() + "'").getPaged()
                .then(response => { 
                    response.results.map(item => {
                        item.teamId = teamId;
                    });
                    dataArray = (dataArray.length === 0) ? response.results : [...dataArray, ...response.results];
                    if (response.hasNext) recurse(response);
                })
                .catch(error => {
                    console.log(error);
                    alert('The Calendar Name your entered: "' + site.Title.toUpperCase() + '" is not valid. Edit that entry or remove it from the list of calendars to avoid any issues with the Schedule Viewer.');
                });

                // Increment counter
                teamId++;
                
            } 
        };

        getCalendarData(calendars)
        .then(_ => {

            $(() => {
                let renamedDataArray = dataArray.map((item) => {
                    item.EndDate = getLocalDate(item.EndDate);
                    item.EventDate = getLocalDate(item.EventDate);
                    
                    // Fix user errors
                    if (item.fRecurrence) item.EndDate = item.EventDate;
                    
                    // Strip out all HTML tags
                    let htmlString = (item.Description != null) ? '<div>' + item.Description + '</div>' : '';
                    item.Description = (htmlString.length > 0) ? $('<div>').html(htmlString).find('body').text() : '';

                    // Other Kendo Scheduler attributes
                    item['recurrenceId'] = null;
                    item['recurrenceException'] = '';
                    item['endTimezone'] = '';
                    item['startTimezone'] = '';

                    // Parse Outlook recurrence data to ical format
                    if (item.RecurrenceData != null && item.RecurrenceData.indexOf('rule') != -1) item.RecurrenceData = recurrenceParser(item.RecurrenceData);

                    // Rename object keys
                    return {
                        id: item.Id,
                        title: item.Title,
                        start: item.EventDate,
                        end: item.EndDate,
                        location: item.Location,
                        description: item.Description,
                        isAllDay: item.fAllDayEvent,
                        fRecurrence: item.fRecurrence,
                        category: item.Category,
                        recurrenceRule: item.RecurrenceData,
                        recurrenceId: item.recurrenceId,
                        recurrenceException: item.recurrenceException,
                        endTimezone: item.endTimezone,
                        startTimezone: item.startTimezone,
                        teamId: item.teamId
                    };
                });

                // Remove duplicate entries from the calendar
                renamedDataArray = renamedDataArray.reduce((acc, curr) => {
                    if (acc.some(x => x.title === curr.title && x.start === curr.start)) return acc;
                    else return [...acc, curr];
                },new Array());
                
                let ds = new kendo.data.SchedulerDataSource({
                    data: renamedDataArray,
                    schema: {
                        model: {
                            id: 'id',
                            fields: {
                                id: { type: 'number' },
                                title: { type: 'string' },
                                start: { type: 'date' },
                                end: { type: 'date' },
                                startTimezone: { type: 'string' },
                                endTimezone: { type: 'string' },
                                description: { type: 'string' },
                                location: { type: 'string' },
                                category: { type: 'string' },
                                fRecurrence: { type: 'boolean' },
                                recurrenceId: { type: 'number' },
                                recurrenceRule: { type: 'string' },
                                recurrenceException: { type: 'string' },
                                isAllDay: { type: 'boolean' },
                            }
                        }
                    }
                });

                // Create Resources DataSource based on property panel data
                let resources = [{field: 'teamId', title: 'Team'}];
                let teamArray = new Array();
                let strBuilder = null;
                $.each(calendars, (index, value) => {
                    teamArray.push({value: index, text: value['Title'], color: value['Color']});

                    switch (index % 2) {
                        case 0:
                            strBuilder += '<tr><td style="width: 50%; padding: 5px;"><input type="checkbox" id="team' + index + '" class="k-checkbox" checked="checked"><label class="k-checkbox-label" style="padding: 5px; font-weight: bold; color: #ffffff; font-weight: bold; background-color: ' + value['Color'] + '" for="team' + index + '">' + value['Title'] + '</label></td>';
                            break;
                        case 1:
                            strBuilder += '<td style="width: 50%; padding: 5px;"><input type="checkbox" id="team' + index + '" class="k-checkbox" checked="checked"><label class="k-checkbox-label" style="padding: 5px; font-weight: bold; color: #ffffff; font-weight: bold; background-color: ' + value['Color'] + '" for="team' + index + '">' + value['Title'] + '</label></td></tr>';
                            break;
                    }
                });
                if (strBuilder.endsWith('</td>')) strBuilder += '<td style="width: 50%; padding: 5px;"></td></tr>';
                $('#filter').append(strBuilder);

                resources[0]['dataSource'] = teamArray;

                $('#calendar').kendoScheduler({
                    date: new Date(),
                    startTime: new Date(kendo.toString(currentDate, 'yyyy/M/d') + ' 07:00 AM'),
                    endTime: new Date(kendo.toString(currentDate, 'yyyy/M/d') + ' 06:00 PM'),
                    editable: false,
                    views: [
                        { type: 'day' },
                        { type: 'workWeek', selected: true },
                        { type: 'month'},
                        { type: 'agenda'},
                    ],
                    toolbar: [ {name: 'search'} ],
                    dataSource: ds,
                    resources: resources
                });

                // Bind event trigger to checkboxes to enable calendar filtering
                $('#filter :checkbox').on('change', e => {
                    let checked = $.map( $('#filter :checked'), i => parseInt($(<HTMLElement>i)[0].id.slice(4)) );
                    let scheduler = $("#calendar").data("kendoScheduler");

                    scheduler.dataSource.filter({
                        operator: item => {
                            return $.inArray(item.teamId, checked) >= 0;
                        }
                    });
                });

                const Events = {
                };
    
                const Utils = {
                };

    
            });

        })
        .catch(err => {
            return false;
        });

        return SPA.instance;
    }
}