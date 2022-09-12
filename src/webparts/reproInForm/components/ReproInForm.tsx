import { BaseComponentContext } from '@microsoft/sp-component-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/taxonomy";
import { ModernTaxonomyPicker } from '@pnp/spfx-controls-react';
import { Dropdown, PrimaryButton, Shimmer } from 'office-ui-fabric-react';

import * as React from 'react';
import { Controller, useForm } from "react-hook-form";

type ReproInFormProps = {
  context: BaseComponentContext
}

type TermId = {
  label: string,
  id: string,
  languageTag: string
}

type FormValues = {
  termSetId: string;
  someTaxoVal: TermId[];
}

type TermSet = {
  id: string,
  name: string,
  groupId: string,
  groupName: string
}


const getAllTermSets = async (context: BaseComponentContext): Promise<TermSet[]> => {
  const sp = spfi().using(SPFx(context));

  const groups = await sp.termStore.groups();
  const sets = await Promise.all(groups.map(async (g) => ({
    g,
    sets: await sp.termStore.groups.getById(g.id).sets()
  })));

  return sets.reduce((acc, current) => {
    current.sets.forEach(ts => {
      acc.push({
        groupId: current.g.id,
        groupName: current.g.name,
        id: ts.id,
        name: ts.localizedNames[0].name
      });
    });
    return acc;
  }, [] as TermSet[]);
}

const ReproInForm = ({ context }: ReproInFormProps): JSX.Element => {


  const defaultValues: FormValues = {
    termSetId: null,
    someTaxoVal: [    ]
  };

  const { handleSubmit, control, getValues, resetField } = useForm({ defaultValues });

  const [data, setData] = React.useState<FormValues | undefined>();
  const [knownTermSets, setKnownTermSets] = React.useState<TermSet[] | undefined>(undefined);

  React.useEffect(() => {
    getAllTermSets(context).then(setKnownTermSets).catch(alert)

  }, [])

  return (
    <>
      <form onSubmit={handleSubmit((data) => setData(data))}>
        <Controller
          name="termSetId"
          control={control}
          render={({ field: { onBlur, onChange, value } }) => {

            return (
              knownTermSets ?
                (
                  <Dropdown
                    defaultValue={value}
                    onChange={(evt, option) => {
                      onChange(option?.key);
                      resetField("someTaxoVal");
                    }}
                    options={knownTermSets.map(ts => ({
                      key: ts.id,
                      text: `[${ts.groupName}] ${ts.name}`
                    }))}
                    onBlur={onBlur}
                    label="Choose term set"
                  />
                )
                : (
                  <Shimmer />
                )
            )
          }}
        />
        <Controller
          name="someTaxoVal"
          control={control}
          render={({ field: { onChange, value, } }) => {
            const termSetId = getValues('termSetId');
            return (
              termSetId ? (
                <ModernTaxonomyPicker
                  termSetId={termSetId}
                  panelTitle={'Choose a term'}
                  label={''}
                  allowMultipleSelections
                  onChange={(selection) => {
                    // Actual output of onChange is not serializable, so wrap it in minimal required value
                    const newVal = selection.map<TermId>(term => ({ id: term.id, label: term.labels[0].name, languageTag: term.labels[0].languageTag }));
                    // Check if term ids have actually changed by comparing IDs
                    const selectedKeys = value.map(s => s.id);
                    const newKeys = selection.map(s => s.id);
                    const isSame = (
                      selectedKeys.length === newKeys.length &&
                      selectedKeys.every(id => newKeys.indexOf(id) !== -1)
                    );
                    if (!isSame) {
                      onChange(newVal);
                    }
                  }}
                  initialValues={
                    value.map(v => ({
                      id: v.id,
                      labels: [{ name: v.label, isDefault: true, languageTag: v.languageTag }]
                    }))
                  }
                  context={context}
                />) : (
                <Shimmer />
              ));
          }
          } />

        <PrimaryButton text='Submit' type='submit' />
      </form>
      <pre>{data && JSON.stringify(data, null, 2)}</pre>
    </>
  );
};


export { ReproInForm };
