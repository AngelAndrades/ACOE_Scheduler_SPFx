import * as $ from 'jquery';
import '@progress/kendo-ui';
import { sp } from '@pnp/sp/presets/all';

interface DataSourceConfig {
    guid: string;
    dsName: string;
    schema: kendo.data.DataSourceSchemaModel;
    pageSize?: number;
    filter?: kendo.data.DataSourceFilter;
    group?: kendo.data.DataSourceGroupItem;
    top?: number;
}

//Object.keys($('#faqGrid').data('kendoGrid').dataSource.options.schema.model.fields)
const cleanseModel = (dataItem: object, spFields: string[]): object => {
    $.each(dataItem, (k,v) => {
        if (spFields.join(',').indexOf(k) == -1) delete dataItem[k];
    });
    return dataItem;
};

export const ds = (args: DataSourceConfig): kendo.data.SchedulerDataSource => {
    var dataArray = [];

    return new kendo.data.SchedulerDataSource({
        transport: {
            create: async options => {
                //console.log('create: ', options.data);
                //console.log('keys: ', Object.keys(args.schema.fields));                
                
                await sp.web.lists.getById(args.guid).items.add(cleanseModel(options.data, Object.keys(args.schema.fields)))
                .then(response => {
                    options.success(response.data);
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to create item');
                });
            },
            read: async options => {
                await sp.web.lists.getById(args.guid).items.select(Object.keys(args.schema.fields).join(',')).top(1000).getPaged()
                .then(response => {
                    const recurse = (next: any) => {
                        next.getNext().then(nestedResponse => {
                            dataArray = [...dataArray, ...nestedResponse.results];
                            if (nestedResponse.hasNext) recurse(nestedResponse);
                            else options.success(dataArray);
                        });
                    };

                    dataArray = response.results;
                    if (response.hasNext) recurse(response);
                    else options.success(dataArray);
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to read items');
                });
            },
            update: async options => {
                await sp.web.lists.getById(args.guid).items.getById(options.data.Id).update(cleanseModel(options.data, Object.keys(args.schema.fields)))
                .then(response => {
                    options.success();
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to update item');
                });
            },
            destroy: async options => {
                await sp.web.lists.getById(args.guid).items.getById(options.data.Id).recycle()
                .then(response => {
                    options.success();
                })
                .catch(error => {
                    console.log(error);
                    throw new Error(args.dsName + ' error, unable to delete item');
                });
            }
        },
        schema: { model: args.schema },
        pageSize: args.pageSize,
        filter: args.filter,
        group: args.group
    });
};