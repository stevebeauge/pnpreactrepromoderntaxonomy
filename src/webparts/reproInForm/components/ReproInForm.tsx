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
  id: string
}

type FormValues = {
  someText: string;
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
        name: ts.description
      });
    });
    return acc;
  }, [] as TermSet[]);
}

const ReproInForm = ({ context }: ReproInFormProps): JSX.Element => {


  const defaultValues: FormValues = {
    someText: null,
    someTaxoVal: []
  };

  const { handleSubmit, control, getValues } = useForm({ defaultValues });

  const [data, setData] = React.useState<FormValues | undefined>();
  const [knownTermSets, setKnownTermSets] = React.useState<TermSet[] | undefined>(undefined);

  React.useEffect(() => {
    getAllTermSets(context).then(setKnownTermSets).catch(alert)

  }, [])

  return (
    <>
      <form onSubmit={handleSubmit((data) => setData(data))}>
        <Controller
          name="someText"
          control={control}
          render={({ field: { onBlur, onChange, value, name } }) => {

            return (
              knownTermSets ?
                (
                  <Dropdown
                    defaultValue={value}
                    onChange={(evt, option) => onChange(option?.key)}
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
          render={({ field: { onBlur, onChange, value, name, }, formState, fieldState }) => {
            const termSetId = getValues('someText');
            return (
              termSetId ? (
                <ModernTaxonomyPicker
                  termSetId={termSetId}
                  panelTitle={'Choose a term'}
                  label={''}
                  onChange={(newVal) => onChange( // Actual output of onChange is not serializable, so wrap it in minimal required value
                    newVal.map<TermId>(term => ({ id: term.id, label: term.labels[0].name }))
                  )}
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
